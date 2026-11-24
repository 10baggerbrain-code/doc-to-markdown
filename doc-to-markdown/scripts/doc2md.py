#!/usr/bin/env python3
"""doc2md 0.2.0-public — 있는 문서를 마크다운으로 바꾼다.

원칙: 원본은 읽기만 / 새 파일에 생성일 / 덮어쓰지 않음(같은 이름이 있으면 -v2, -v3)
      / 같은 원본 해시면 다시 변환하지 않음(멱등) / 기록은 해시 사슬로 덧붙이기만
      / 개인정보 모양은 개수만 알리고 값은 출력하지 않음 / 확인 못 한 것은 Unverified.
사용: python3 doc2md.py <파일이나 폴더> [...] --out <출력 폴더>
"""
import argparse, datetime, hashlib, json, os, re, shutil, struct, subprocess, sys, zipfile, zlib
from pathlib import Path

VERSION = "0.2.0-public"
TODAY = datetime.date.today().isoformat()
SUPPORTED = {".docx", ".pptx", ".xlsx", ".xlsm", ".csv", ".pdf", ".html", ".htm",
             ".hwpx", ".hwp", ".txt", ".md", ".rtf", ".odt", ".epub", ".json"}
NEEDS_OTHER = {".doc": "DOC(옛 Word)는 LibreOffice로 DOCX 변환이 필요",
               ".ppt": "PPT(옛 PowerPoint)는 LibreOffice로 PPTX 변환이 필요",
               ".xls": "XLS(옛 Excel)는 LibreOffice로 XLSX 변환이 필요"}
# 문서 속 글은 AI에 그대로 들어간다 → AI에게 하는 지시문 모양은 개수만 알린다
INJECT = r"(?i)(ignore (all |any )?(previous|prior|above) (instructions|prompts)|system prompt|you are now|disregard .{0,20}instructions|이전 (지시|지침|명령)(을|를)? (모두 )?무시|이 문서를 읽는 ?(AI|에이전트|모델)|(너는|당신은) 이제 |시스템 프롬프트)"
# 받은 폴더 안 AI 지침 파일은 변환하지 않는다(변환본을 통해 남의 지침이 내 AI에 들어오지 않게)
AGENT_FILES = {"claude.md", "agents.md", "gemini.md", ".cursorrules", "copilot-instructions.md"}
AGENT_DIRS = {".claude", ".cursor", ".gemini", ".codex"}
PII = {"email": r"[\w.+-]+@[\w-]+\.[\w.-]+",
       "phone_kr": r"01[016-9][-. ]?\d{3,4}[-. ]?\d{4}",
       "rrn_kr": r"\b\d{6}[- ]?[1-8]\d{6}\b"}


def sha256(p):
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def run(cmd):
    r = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
    if r.returncode != 0:
        raise RuntimeError(r.stderr.strip()[:300] or "실패")
    return r.stdout


def conv_pandoc(p, fmt=None):
    cmd = ["pandoc", str(p), "-t", "gfm", "--wrap=none"]
    if fmt:
        cmd[2:2] = ["-f", fmt]
    return run(cmd), "pandoc " + run(["pandoc", "--version"]).split("\n")[0].split()[-1]


def conv_pdf(p):
    notes = []
    try:
        import pdfplumber
        parts = []
        with pdfplumber.open(p) as pdf:
            for i, page in enumerate(pdf.pages, 1):
                t = page.extract_text() or ""
                parts.append(f"\n\n<!-- 쪽 {i} -->\n\n{t.strip()}")
                for tb in page.extract_tables() or []:
                    parts.append(table_md(tb))
        text, tool = "".join(parts), "pdfplumber"
    except Exception as e:  # 예비: pdftotext
        text, tool = run(["pdftotext", "-layout", str(p), "-"]), "pdftotext"
        notes.append(f"pdfplumber 실패로 pdftotext 사용({e.__class__.__name__})")
    if len(re.sub(r"\s|<!--.*?-->", "", text)) < 20:
        notes.append("글자가 거의 없음: 스캔 PDF일 수 있음. OCR 필요(이 변환은 내용 없음)")
    return text, tool, notes


def table_md(rows):
    rows = [[("" if c is None else str(c)).replace("\n", " ").replace("|", "\\|") for c in r] for r in rows if r]
    if not rows:
        return ""
    w = max(len(r) for r in rows)
    rows = [r + [""] * (w - len(r)) for r in rows]
    out = ["| " + " | ".join(rows[0]) + " |", "|" + " --- |" * w]
    out += ["| " + " | ".join(r) + " |" for r in rows[1:]]
    return "\n\n" + "\n".join(out) + "\n"


def conv_xlsx(p):
    import openpyxl
    wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
    parts, notes = [], []
    for ws in wb.worksheets:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        rows = [r for r in rows if any(c is not None for c in r)]
        parts.append(f"\n\n## 시트: {ws.title}\n" + (table_md(rows) if rows else "\n(빈 시트)\n"))
    notes.append("수식은 마지막 저장된 값으로 옮김(수식 자체는 옮기지 않음)")
    wf = openpyxl.load_workbook(p, read_only=True)
    lost = sum(1 for ws1, ws2 in zip(wf.worksheets, wb.worksheets)
               for r1, r2 in zip(ws1.iter_rows(values_only=True), ws2.iter_rows(values_only=True))
               for c1, c2 in zip(r1, r2) if isinstance(c1, str) and c1.startswith("=") and c2 is None)
    if lost:
        notes.append(f"저장된 값이 없는 수식 칸 {lost}개는 빈칸으로 옮겨짐. Excel에서 한 번 열어 저장한 뒤 다시")
    return "".join(parts), "openpyxl", notes


def conv_pptx(p):
    from pptx import Presentation
    prs, parts = Presentation(p), []
    for i, s in enumerate(prs.slides, 1):
        parts.append(f"\n\n## 슬라이드 {i}\n")
        for sh in s.shapes:
            if sh.has_text_frame:
                for para in sh.text_frame.paragraphs:
                    t = "".join(r.text for r in para.runs).strip()
                    if t:
                        parts.append(f"\n{'  ' * para.level}- {t}")
            if getattr(sh, "has_table", False) and sh.has_table:
                parts.append(table_md([[c.text for c in r.cells] for r in sh.table.rows]))
        if s.has_notes_slide and s.notes_slide.notes_text_frame.text.strip():
            parts.append("\n\n> 발표자 메모: " + s.notes_slide.notes_text_frame.text.strip().replace("\n", " "))
    return "".join(parts), "python-pptx", ["그림·도형 배치는 옮기지 않음(글자만)"]


def conv_hwpx(p):
    with zipfile.ZipFile(p) as z:
        secs = sorted(n for n in z.namelist() if re.match(r"Contents/section\d+\.xml", n))
        parts = []
        for n in secs:
            xml = z.read(n).decode("utf-8", "replace")
            for para in re.findall(r"<hp:p[ >].*?</hp:p>", xml, flags=re.S):
                t = "".join(re.findall(r"<hp:t[^>]*>(.*?)</hp:t>", para, flags=re.S))
                t = re.sub(r"<[^>]+>", "", t).strip()
                if t:
                    parts.append(t)
    import html
    return "\n\n".join(html.unescape(x) for x in parts), "hwpx-zip-xml", ["HWPX는 문단 글자만 옮김(표는 칸 글자만 차례로, 표 모양·서식 Unverified)"]


# --- HWP 5.0 읽기 (한컴 공개 형식 문서 기준, 표준 라이브러리만) ---------------------------
# 한글과컴퓨터가 공개한 HWP 5.0 형식 문서를 참고해 직접 작성. 외부 HWP 도구의 코드는 쓰지 않음.
# 본 제품은 한글과컴퓨터의 한/글 문서 파일(.hwp) 공개 문서를 참고하여 개발하였습니다.

CFB_SIG = bytes.fromhex("D0CF11E0A1B11AE1")
FREE, ENDCHAIN = 0xFFFFFFFF, 0xFFFFFFFE


class CFB:
    """복합 문서(OLE) 읽기. 읽기만 하고, 사슬이 돌거나 범위를 넘으면 멈춘다."""

    def __init__(self, data):
        if data[:8] != CFB_SIG:
            raise ValueError("OLE 복합 문서가 아님")
        self.d = data
        self.ss = 1 << struct.unpack_from("<H", data, 0x1E)[0]
        self.mss = 1 << struct.unpack_from("<H", data, 0x20)[0]
        n_fat, first_dir = struct.unpack_from("<II", data, 0x2C)
        self.cutoff, first_mf, n_mf, first_difat, n_difat = struct.unpack_from("<IIIII", data, 0x38)
        difat = list(struct.unpack_from("<109I", data, 0x4C))
        s, guard = first_difat, 0
        while s not in (FREE, ENDCHAIN) and guard < n_difat + 1:
            ent = struct.unpack_from(f"<{self.ss // 4}I", self._sec(s))
            difat += ent[:-1]
            s, guard = ent[-1], guard + 1
        fat = b"".join(self._sec(x) for x in difat[:n_fat] if x not in (FREE, ENDCHAIN))
        self.fat = struct.unpack(f"<{len(fat) // 4}I", fat)
        dirs = self._chain(first_dir, self.fat, self._sec)
        self.entries = []
        for i in range(0, len(dirs) - 127, 128):
            e = dirs[i:i + 128]
            nlen = struct.unpack_from("<H", e, 0x40)[0]
            name = e[:max(0, nlen - 2)].decode("utf-16-le", "replace")
            typ = e[0x42]
            left, right, child = struct.unpack_from("<III", e, 0x44)
            start, size = struct.unpack_from("<II", e, 0x74)
            self.entries.append(dict(name=name, type=typ, left=left, right=right, child=child, start=start, size=size))
        root = self.entries[0]
        mf = self._chain(first_mf, self.fat, self._sec) if n_mf else b""
        self.minifat = struct.unpack(f"<{len(mf) // 4}I", mf) if mf else ()
        self.ministream = self._chain(root["start"], self.fat, self._sec)[:root["size"]] if root["size"] else b""
        self.paths = {}
        self._walk(root["child"], "", set())

    def _sec(self, n):
        off = (n + 1) * self.ss
        if off + self.ss > len(self.d):
            raise ValueError("섹터가 파일 밖을 가리킴")
        return self.d[off:off + self.ss]

    def _minisec(self, n):
        off = n * self.mss
        return self.ministream[off:off + self.mss]

    def _chain(self, s, table, reader):
        out, seen = [], set()
        while s not in (FREE, ENDCHAIN):
            if s in seen or s >= len(table):
                raise ValueError("사슬이 돌거나 범위를 넘음")
            seen.add(s)
            out.append(reader(s))
            s = table[s]
        return b"".join(out)

    def _walk(self, i, prefix, seen):
        if i == FREE or i in seen or i >= len(self.entries):
            return
        seen.add(i)
        e = self.entries[i]
        self._walk(e["left"], prefix, seen)
        self._walk(e["right"], prefix, seen)
        path = prefix + e["name"]
        self.paths[path] = e
        if e["type"] == 1:
            self._walk(e["child"], path + "/", seen)

    def read(self, path):
        e = self.paths[path]
        if e["size"] < self.cutoff:
            return self._chain(e["start"], self.minifat, self._minisec)[:e["size"]]
        return self._chain(e["start"], self.fat, self._sec)[:e["size"]]


PARA_TEXT = 0x10 + 51
CHAR_CTRL = {0, 10, 13, 24, 25, 26, 27, 28, 29, 30, 31}  # 1글자 크기, 나머지 0~31은 8글자 크기


def para_text(b):
    w = struct.unpack(f"<{len(b) // 2}H", b[:len(b) // 2 * 2])
    out, i = [], 0
    while i < len(w):
        c = w[i]
        if c >= 32:
            out.append(chr(c))
            i += 1
        elif c in CHAR_CTRL:
            out.append({10: "\n", 24: "-", 30: " ", 31: " "}.get(c, ""))
            i += 1
        else:
            if c == 9:
                out.append("\t")
            i += 8
    s = "".join(out)
    try:  # 서로게이트 쌍(이모지 등) 합치기
        s = s.encode("utf-16", "surrogatepass").decode("utf-16")
    except UnicodeError:
        pass
    return s


def records(b):
    i = 0
    while i + 4 <= len(b):
        h = struct.unpack_from("<I", b, i)[0]
        tag, size = h & 0x3FF, (h >> 20) & 0xFFF
        i += 4
        if size == 0xFFF:
            size = struct.unpack_from("<I", b, i)[0]
            i += 4
        if i + size > len(b):
            raise ValueError("레코드가 스트림 밖으로 넘음")
        yield tag, b[i:i + size]
        i += size


def read_hwp5(data):
    """돌려줌: (본문 글, 미리보기 글, 알림 목록). 암호·배포용 문서는 ValueError."""
    try:
        f = CFB(data)
    except (IndexError, struct.error) as e:
        raise ValueError(f"OLE 구조가 깨짐({e.__class__.__name__})")
    if "FileHeader" not in f.paths:
        raise ValueError("HWP FileHeader 없음(HWP 5.0이 아님)")
    fh = f.read("FileHeader")
    if not fh.startswith(b"HWP Document File"):
        raise ValueError("HWP 서명이 아님")
    ver = struct.unpack_from("<I", fh, 32)[0]
    prop = struct.unpack_from("<I", fh, 36)[0]
    if prop & 2:
        raise ValueError("암호가 걸린 문서. 한글에서 암호를 풀고 다른 이름으로 저장한 뒤 다시")
    if prop & 4:
        raise ValueError("배포용 문서(본문 암호화). 한글에서 배포용 해제 후 다른 이름으로 저장한 뒤 다시")
    secs = sorted((p for p in f.paths if p.startswith("BodyText/Section")),
                  key=lambda p: int(p.rsplit("Section", 1)[1] or 0))
    if not secs:
        raise ValueError("BodyText 구역이 없음")
    paras = []
    for p in secs:
        raw = f.read(p)
        if prop & 1:
            raw = zlib.decompress(raw, -15)
        for tag, body in records(raw):
            if tag == PARA_TEXT:
                t = para_text(body).strip()
                if t:
                    paras.append(t)
    prv = ""
    if "PrvText" in f.paths:
        prv = f.read("PrvText").decode("utf-16-le", "replace")
    v = f"{ver >> 24}.{(ver >> 16) & 255}.{(ver >> 8) & 255}.{ver & 255}"
    notes = [f"HWP {v}, 구역 {len(secs)}개, 표는 칸 글자만 차례로 옮김(표 모양 Unverified)"]
    return "\n\n".join(paras), prv, notes


def prv_match(body, prv):
    """미리보기 글의 낱말 가운데 본문에 있는 비율. 미리보기가 없으면 None."""
    import re
    words = [w for w in re.findall(r"[0-9A-Za-z가-힣]{2,}", re.sub(r"<[^>]*>", " ", prv))]
    if not words:
        return None
    bset = set(re.findall(r"[0-9A-Za-z가-힣]{2,}", body))
    return sum(1 for w in words if w in bset) / len(words)


def conv_hwp(p):
    data = p.read_bytes()
    if data[:8] == CFB_SIG:
        body, prv, notes = read_hwp5(data)
        r = prv_match(body, prv)
        if r is None:
            notes.append("미리보기 글(PrvText)이 없어 대조 못 함")
        else:
            notes.append(f"미리보기 글 대조 {r:.0%}" + ("" if r >= 0.9 else " — 90% 미만, 원본과 나란히 확인 필요"))
        return body, "hwp5-공개문서-파서", notes
    if data[:20].startswith(b"HWP Document File V3"):
        import tempfile
        with tempfile.TemporaryDirectory() as td:
            run(["soffice", "--headless", f"-env:UserInstallation=file://{td}/u", "--convert-to", "docx", "--outdir", td, str(p)])
            out = next(Path(td).glob("*.docx"), None)
            if not out:
                raise RuntimeError("LibreOffice가 HWP 3.x를 바꾸지 못함")
            t, _ = conv_pandoc(out)
        return t, "libreoffice-hwp3+pandoc", ["HWP 3.x는 LibreOffice 필터로 바꿈(정확도 Unverified, 원본과 나란히 확인 필요)"]
    raise ValueError("HWP 서명을 알아보지 못함(HWP 5.0·3.x 아님)")


def convert(p):
    ext = p.suffix.lower()
    if ext == ".pdf":
        return conv_pdf(p)
    if ext in (".xlsx", ".xlsm"):
        return conv_xlsx(p)
    if ext == ".pptx":
        return conv_pptx(p)
    if ext == ".hwpx":
        return conv_hwpx(p)
    if ext == ".hwp":
        return conv_hwp(p)
    if ext == ".csv":
        import csv
        with open(p, newline="", encoding="utf-8-sig", errors="replace") as f:
            return table_md(list(csv.reader(f))), "csv", []
    if ext in (".txt", ".md", ".json"):
        t = p.read_text(encoding="utf-8", errors="replace")
        return (f"```json\n{t}\n```" if ext == ".json" else t), "copy", []
    fmt = {".htm": "html"}.get(ext)
    t, tool = conv_pandoc(p, fmt)
    return t, tool, []


def free_name(out_dir, stem):
    base = out_dir / f"{stem}-{TODAY}.md"
    n = 2
    while base.exists():
        base = out_dir / f"{stem}-{TODAY}-v{n}.md"
        n += 1
    return base


def load_ledger(path):
    rows = []
    if path.exists():
        rows = [json.loads(l) for l in path.read_text(encoding="utf-8").splitlines() if l.strip()]
    return rows


def verify_chain(rows):
    prev = "0" * 64
    for r in rows:
        body = {k: v for k, v in r.items() if k != "hash"}
        if body.get("prev") != prev or r["hash"] != hashlib.sha256(json.dumps(body, ensure_ascii=False, sort_keys=True).encode()).hexdigest():
            return False
        prev = r["hash"]
    return True


def append(path, rows, entry):
    entry["prev"] = rows[-1]["hash"] if rows else "0" * 64
    entry["hash"] = hashlib.sha256(json.dumps(entry, ensure_ascii=False, sort_keys=True).encode()).hexdigest()
    with open(path, "a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=False, sort_keys=True) + "\n")
    rows.append(entry)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("inputs", nargs="+")
    ap.add_argument("--out", required=True)
    ap.add_argument("--date", help="사용자 현지 날짜 YYYY-MM-DD(서버 시계는 UTC라 하루 어긋날 수 있음)")
    a = ap.parse_args()
    global TODAY
    if a.date:
        TODAY = datetime.date.fromisoformat(a.date).isoformat()
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    ledger_path = out / "doc2md-ledger.jsonl"
    rows = load_ledger(ledger_path)
    if not verify_chain(rows):
        print("멈춤: 기록 파일의 해시 사슬이 맞지 않음. 고치지 말고 사용자에게 알릴 것.", file=sys.stderr)
        sys.exit(2)
    done = {r["src_sha256"]: r for r in rows if r.get("status") == "converted"}
    failed_seen = {(r["src_sha256"], r.get("error")) for r in rows if r.get("status") == "failed"}
    files = []
    for i in a.inputs:
        q = Path(i)
        files += sorted(x for x in q.rglob("*") if x.is_file()) if q.is_dir() else [q]
    out_abs = out.resolve()
    report = {"converted": 0, "skipped_same_hash": 0, "unsupported": 0, "failed": 0, "pii_files": 0, "empty": 0,
              "agent_files_skipped": 0, "inject_like_files": 0}
    for p in files:
        if out_abs in p.resolve().parents:  # 자기 출력은 다시 넣지 않음
            continue
        ext = p.suffix.lower()
        rel = p.parts
        if p.name.lower() in AGENT_FILES or any(x.lower() in AGENT_DIRS for x in rel[:-1]):
            print(f"[Confirmed] 건너뜀 {p.name}: AI 지침 파일·폴더는 변환하지 않음(데이터로만 둠)")
            report["agent_files_skipped"] += 1
            continue
        if ext not in SUPPORTED:
            why = NEEDS_OTHER.get(ext, "지원하지 않는 형식")
            print(f"[Unverified] 건너뜀 {p.name}: {why}")
            report["unsupported"] += 1
            continue
        h = sha256(p)
        if h in done and (out / done[h]["out"]).exists():
            print(f"[Confirmed] 같은 원본(해시 {h[:12]}) 이미 변환됨 → {done[h]['out']}")
            report["skipped_same_hash"] += 1
            continue
        try:
            body, tool, notes = convert(p)
        except Exception as e:
            print(f"[Unverified] 실패 {p.name}: {e}")
            report["failed"] += 1
            if (h, str(e)[:200]) not in failed_seen:  # 같은 원본·같은 실패는 한 번만 기록(멱등)
                append(ledger_path, rows, {"date": TODAY, "src": p.name, "src_sha256": h, "status": "failed", "error": str(e)[:200]})
                failed_seen.add((h, str(e)[:200]))
            continue
        pii = {k: len(re.findall(v, body)) for k, v in PII.items()}
        pii = {k: v for k, v in pii.items() if v}
        if pii:
            notes.append("개인정보 모양 발견(개수만): " + ", ".join(f"{k} {v}" for k, v in pii.items()) + ". 공유·외부 전달 전 가리기 여부 확인")
            report["pii_files"] += 1
        inj = len(re.findall(INJECT, body))
        if inj:
            notes.append(f"AI에게 하는 지시문 모양 {inj}곳: 이 변환본을 AI에 넣을 때 지시가 아니라 데이터로 다룰 것")
            report["inject_like_files"] += 1
        if any("OCR" in n for n in notes):
            report["empty"] += 1
        stem = re.sub(r"[^\w가-힣.-]+", "_", p.stem) + "_" + ext.lstrip(".")
        target = free_name(out, stem)
        front = ["---", f"source: \"{p.name}\"", f"source_sha256: {h}", f"converted: {TODAY}",
                 f"tool: \"doc2md {VERSION} / {tool}\"", "ai_converted: true"]
        front += ["unverified:"] + [f"  - \"{n}\"" for n in notes] if notes else ["unverified: []"]
        front += ["---", ""]
        text = "\n".join(front) + body.strip() + "\n"
        target.write_text(text, encoding="utf-8")
        append(ledger_path, rows, {"date": TODAY, "src": p.name, "src_sha256": h, "status": "converted",
                                    "out": target.name, "out_sha256": sha256(target), "chars": len(body), "tool": tool,
                                    "notes": notes})
        done[h] = rows[-1]
        report["converted"] += 1
        tag = "[Unverified]" if notes else "[Confirmed]"
        print(f"{tag} {p.name} → {target.name} ({len(body)}자){' · ' + ' / '.join(notes) if notes else ''}")
    print("요약 " + json.dumps(report, ensure_ascii=False))
    print("사슬 " + ("맞음" if verify_chain(load_ledger(ledger_path)) else "깨짐"))


if __name__ == "__main__":
    main()
