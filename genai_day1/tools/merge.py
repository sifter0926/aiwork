"""pptxgenjs로 만든 새 슬라이드를 원본 덱의 지정 위치 뒤에 이식한다.
usage: merge.py orig_unpacked_dir new_unpacked_dir insert_after(1-based) out_dir
"""
import re
import shutil
import sys
from pathlib import Path

orig, new, after, out = Path(sys.argv[1]), Path(sys.argv[2]), int(sys.argv[3]), Path(sys.argv[4])
if out.exists():
    shutil.rmtree(out)
shutil.copytree(orig, out)

R_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
R_NOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
R_LAYOUT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"


def rels_of(p: Path) -> Path:
    return p.parent / "_rels" / (p.name + ".rels")


def max_num(folder: Path, stem: str) -> int:
    nums = [int(m.group(1)) for f in folder.glob(f"{stem}*.xml") if (m := re.match(rf"{stem}(\d+)\.xml", f.name))]
    return max(nums, default=0)


# 새 덱의 슬라이드 순서
npres = (new / "ppt/presentation.xml").read_text(encoding="utf-8")
nrels = (new / "ppt/_rels/presentation.xml.rels").read_text(encoding="utf-8")
rid2t = {m.group(1): m.group(2) for m in re.finditer(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', nrels)}
rid2t.update({m.group(2): m.group(1) for m in re.finditer(r'Target="([^"]+)"[^>]*Id="(rId\d+)"', nrels)})
new_slides = [rid2t[r].split("/")[-1] for r in re.findall(r'<p:sldId [^>]*r:id="(rId\d+)"', npres)]

(out / "ppt/charts/_rels").mkdir(parents=True, exist_ok=True)
(out / "ppt/embeddings").mkdir(parents=True, exist_ok=True)
slide_n = max_num(out / "ppt/slides", "slide")
notes_n = max_num(out / "ppt/notesSlides", "notesSlide")
chart_n = max_num(out / "ppt/charts", "chart")
emb_n = len(list((out / "ppt/embeddings").glob("*.xlsx")))

overrides, added = [], []
TRENDLINE = ('<c:trendline><c:spPr><a:ln w="22225" cap="rnd"><a:solidFill><a:srgbClr val="A64F16"/></a:solidFill>'
             '<a:prstDash val="dash"/></a:ln></c:spPr><c:trendlineType val="linear"/><c:dispRSqr val="1"/><c:dispEq val="1"/>'
             '<c:trendlineLbl><c:layout><c:manualLayout><c:x val="-0.02"/><c:y val="0.42"/></c:manualLayout></c:layout>'
             '<c:numFmt formatCode="0.00" sourceLinked="0"/><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>'
             '<a:defRPr sz="1100" b="1"><a:solidFill><a:srgbClr val="A64F16"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr>'
             '</a:pPr><a:endParaRPr lang="ko-KR"/></a:p></c:txPr></c:trendlineLbl></c:trendline>')

for name in new_slides:
    src = new / "ppt/slides" / name
    slide_n += 1
    dst_name = f"slide{slide_n}.xml"
    xml = src.read_text(encoding="utf-8")
    # 제목 텍스트 상자를 원본처럼 제목 개체 틀로 (개요 보기·접근성)
    xml = re.sub(r'(<p:cNvPr id="\d+" name="Title"></p:cNvPr><p:cNvSpPr) txBox="1"(/><p:nvPr>)(</p:nvPr>)',
                 r'\1\2<p:ph type="title"/>\3', xml)
    xml = xml.replace('<a:ea typeface="Arial"', '<a:ea typeface="맑은 고딕"')  # 원본 PDF처럼 한글은 맑은 고딕
    xml = re.sub(r'<p:cSld name="[^"]*">', f'<p:cSld name="Slide {slide_n}">', xml)
    (out / "ppt/slides" / dst_name).write_text(xml, encoding="utf-8")
    rels = rels_of(src).read_text(encoding="utf-8")

    def fix(m):
        global notes_n, chart_n, emb_n
        rid, typ, tgt = m.group(1), m.group(2), m.group(3)
        if typ == R_LAYOUT:
            tgt = "../slideLayouts/slideLayout2.xml"
        elif typ == R_NOTES:
            notes_n += 1
            nsrc = new / "ppt/notesSlides" / tgt.split("/")[-1]
            nname = f"notesSlide{notes_n}.xml"
            shutil.copy(nsrc, out / "ppt/notesSlides" / nname)
            nr = rels_of(nsrc).read_text(encoding="utf-8")
            nr = re.sub(r'Target="[^"]*slides/slide\d+\.xml"', f'Target="../slides/{dst_name}"', nr)
            nr = re.sub(r'Target="[^"]*notesMasters/notesMaster\d+\.xml"', 'Target="../notesMasters/notesMaster1.xml"', nr)
            rels_of(out / "ppt/notesSlides" / nname).write_text(nr, encoding="utf-8")
            overrides.append((f"/ppt/notesSlides/{nname}", "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"))
            tgt = f"../notesSlides/{nname}"
        elif typ == R_CHART:
            chart_n += 1
            csrc = new / "ppt/charts" / tgt.split("/")[-1]
            cname = f"chart{chart_n}.xml"
            cx = csrc.read_text(encoding="utf-8")
            if "<c:scatterChart>" in cx:  # 엑셀의 '추세선 추가 > 선형 · 수식 · R²'와 같은 네이티브 추세선
                tl = TRENDLINE
                if "산점도와 추세선" in cx:  # 보고서 안의 작은 차트는 선만 (수치는 옆 표에)
                    tl = re.sub(r'<c:dispRSqr val="1"/><c:dispEq val="1"/><c:trendlineLbl>.*?</c:trendlineLbl>',
                                '<c:dispRSqr val="0"/><c:dispEq val="0"/>', tl)
                cx = cx.replace("<c:xVal>", tl + "<c:xVal>", 1)
            (out / "ppt/charts" / cname).write_text(cx, encoding="utf-8")
            cr = rels_of(csrc).read_text(encoding="utf-8")
            for em in re.findall(r'Target="\.\./embeddings/([^"]+)"', cr):
                emb_n += 1
                ename = f"Microsoft_Excel_Worksheet{emb_n}.xlsx"
                shutil.copy(new / "ppt/embeddings" / em, out / "ppt/embeddings" / ename)
                cr = cr.replace(f"../embeddings/{em}", f"../embeddings/{ename}")
            rels_of(out / "ppt/charts" / cname).write_text(cr, encoding="utf-8")
            overrides.append((f"/ppt/charts/{cname}", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"))
            tgt = f"../charts/{cname}"
        return f'<Relationship Id="{rid}" Type="{typ}" Target="{tgt}"/>'

    rels = re.sub(r'<Relationship Id="(rId\d+)" Type="([^"]+)" Target="([^"]+)"/>', fix, rels)
    rels_of(out / "ppt/slides" / dst_name).write_text(rels, encoding="utf-8")
    overrides.append((f"/ppt/slides/{dst_name}", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"))
    added.append(dst_name)

# Content types
ct_p = out / "[Content_Types].xml"
ct = ct_p.read_text(encoding="utf-8")
ct = ct.replace("</Types>", "".join(f'<Override PartName="{p}" ContentType="{t}"/>' for p, t in overrides) + "</Types>")
ct_p.write_text(ct, encoding="utf-8")

# presentation.xml + rels
pres_p, prels_p = out / "ppt/presentation.xml", out / "ppt/_rels/presentation.xml.rels"
pres, prels = pres_p.read_text(encoding="utf-8"), prels_p.read_text(encoding="utf-8")
max_id = max(int(i) for i in re.findall(r'<p:sldId id="(\d+)"', pres))
max_rid = max(int(i) for i in re.findall(r'Id="rId(\d+)"', prels))
entries, rel_add = [], []
for k, name in enumerate(added, 1):
    rid = f"rId{max_rid + k}"
    entries.append(f'<p:sldId id="{max_id + k}" r:id="{rid}"/>')
    rel_add.append(f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/{name}"/>')
sld = re.findall(r'<p:sldId [^>]*/>', pres)
anchor = sld[after - 1]
pres = pres.replace(anchor, anchor + "".join(entries), 1)
pres_p.write_text(pres, encoding="utf-8")
prels_p.write_text(prels.replace("</Relationships>", "".join(rel_add) + "</Relationships>"), encoding="utf-8")

# docProps/app.xml 슬라이드 수
app_p = out / "docProps/app.xml"
app = app_p.read_text(encoding="utf-8")
total = len(sld) + len(added)
app = re.sub(r"<Slides>\d+</Slides>", f"<Slides>{total}</Slides>", app)
app = re.sub(r"<Notes>\d+</Notes>", f"<Notes>{total}</Notes>", app)
app_p.write_text(app, encoding="utf-8")
print("added", added, "total", total)
