from fastapi import FastAPI, Form
from fastapi.responses import HTMLResponse, StreamingResponse
import io
import zipfile

app = FastAPI()

INDEX_HTML = """
<!DOCTYPE html>
<html>
<head><title>Document Generator</title></head>
<body>
<h1>Generador de Informes y Presentaciones</h1>
<h2>Informe</h2>
<form action='/generate-report' method='post'>
  Título: <input type='text' name='title'><br>
  Contenido:<br>
  <textarea name='content' rows='10' cols='50'></textarea><br>
  <input type='submit' value='Crear Informe'>
</form>

<h2>PPT</h2>
<form action='/generate-ppt' method='post'>
  Título: <input type='text' name='title'><br>
  Contenido:<br>
  <textarea name='content' rows='10' cols='50'></textarea><br>
  <input type='submit' value='Crear PPT'>
</form>
</body>
</html>
"""

@app.get("/", response_class=HTMLResponse)
async def index():
    return INDEX_HTML

def make_report_html(title: str, content: str) -> str:
    pages = []
    for i in range(1, 31):
        pages.append(f"<div class='page'><h2>{title} - Pagina {i}</h2><p>{content}</p></div>")
    html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset='utf-8'>
<title>{title}</title>
<style>
  .page {{ page-break-after: always; }}
</style>
</head>
<body>
{''.join(pages)}
</body>
</html>"""
    return html

@app.post("/generate-report")
async def generate_report(title: str = Form(...), content: str = Form(...)):
    html = make_report_html(title, content)
    return HTMLResponse(html, media_type='text/html')

# PPTX templates
CONTENT_TYPES = """<?xml version='1.0' encoding='UTF-8'?>
<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>
 <Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/>
 <Default Extension='xml' ContentType='application/xml'/>
 <Override PartName='/ppt/presentation.xml' ContentType='application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'/>
 <Override PartName='/ppt/slides/slide1.xml' ContentType='application/vnd.openxmlformats-officedocument.presentationml.slide+xml'/>
 <Override PartName='/ppt/theme/theme1.xml' ContentType='application/vnd.openxmlformats-officedocument.theme+xml'/>
 <Override PartName='/ppt/slideLayouts/slideLayout1.xml' ContentType='application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'/>
 <Override PartName='/ppt/slideMasters/slideMaster1.xml' ContentType='application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'/>
</Types>"""

RELS = """<?xml version='1.0' encoding='UTF-8'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>
 <Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='ppt/presentation.xml'/>
</Relationships>"""

PRESENTATION_RELS = """<?xml version='1.0' encoding='UTF-8'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>
 <Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster' Target='slideMasters/slideMaster1.xml'/>
 <Relationship Id='rId2' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide' Target='slides/slide1.xml'/>
</Relationships>"""

PRESENTATION_XML = """<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<p:presentation xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>
 <p:sldMasterIdLst>
  <p:sldMasterId id='2147483648' r:id='rId1'/>
 </p:sldMasterIdLst>
 <p:sldIdLst>
  <p:sldId id='256' r:id='rId2'/>
 </p:sldIdLst>
 <p:sldSz cx='9144000' cy='6858000' type='screen4x3'/>
</p:presentation>"""

SLIDE1_RELS = """<?xml version='1.0' encoding='UTF-8'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>
 <Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout' Target='../slideLayouts/slideLayout1.xml'/>
</Relationships>"""

SLIDE1_XML_TEMPLATE = """<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<p:sld xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main'>
 <p:cSld>
  <p:spTree>
   <p:nvGrpSpPr>
    <p:cNvPr id='1' name=''/>
    <p:cNvGrpSpPr/>
    <p:nvPr/>
   </p:nvGrpSpPr>
   <p:grpSpPr>
    <a:xfrm>
     <a:off x='0' y='0'/>
     <a:ext cx='0' cy='0'/>
    </a:xfrm>
   </p:grpSpPr>
   <p:sp>
    <p:nvSpPr>
     <p:cNvPr id='2' name='Title'/>
     <p:cNvSpPr/>
     <p:nvPr/>
    </p:nvSpPr>
    <p:spPr/>
    <p:txBody>
     <a:bodyPr/>
     <a:lstStyle/>
     <a:p><a:r><a:t>{title}</a:t></a:r></a:p>
    </p:txBody>
   </p:sp>
   <p:sp>
    <p:nvSpPr>
     <p:cNvPr id='3' name='Content'/>
     <p:cNvSpPr/>
     <p:nvPr/>
    </p:nvSpPr>
    <p:spPr/>
    <p:txBody>
     <a:bodyPr/>
     <a:lstStyle/>
     <a:p><a:r><a:t>{content}</a:t></a:r></a:p>
    </p:txBody>
   </p:sp>
  </p:spTree>
 </p:cSld>
 <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>"""

SLIDE_LAYOUT_XML = """<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<p:sldLayout xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main' type='title'>
 <p:cSld>
  <p:spTree>
   <p:nvGrpSpPr>
    <p:cNvPr id='1' name=''/>
    <p:cNvGrpSpPr/>
    <p:nvPr/>
   </p:nvGrpSpPr>
   <p:grpSpPr>
    <a:xfrm>
     <a:off x='0' y='0'/>
     <a:ext cx='0' cy='0'/>
    </a:xfrm>
   </p:grpSpPr>
  </p:spTree>
 </p:cSld>
 <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>"""

SLIDE_LAYOUT_RELS = """<?xml version='1.0' encoding='UTF-8'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>
 <Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster' Target='../slideMasters/slideMaster1.xml'/>
</Relationships>"""

SLIDE_MASTER_XML = """<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<p:sldMaster xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main'>
 <p:cSld>
  <p:spTree>
   <p:nvGrpSpPr>
    <p:cNvPr id='1' name=''/>
    <p:cNvGrpSpPr/>
    <p:nvPr/>
   </p:nvGrpSpPr>
   <p:grpSpPr>
    <a:xfrm>
     <a:off x='0' y='0'/>
     <a:ext cx='0' cy='0'/>
    </a:xfrm>
   </p:grpSpPr>
  </p:spTree>
 </p:cSld>
 <p:sldLayoutIdLst>
  <p:sldLayoutId id='1' r:id='rId1'/>
 </p:sldLayoutIdLst>
 <p:txStyles/>
</p:sldMaster>"""

SLIDE_MASTER_RELS = """<?xml version='1.0' encoding='UTF-8'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>
 <Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout' Target='../slideLayouts/slideLayout1.xml'/>
 <Relationship Id='rId2' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme' Target='../theme/theme1.xml'/>
</Relationships>"""

THEME1_XML = """<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<a:theme xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main' name='Custom Theme'>
 <a:themeElements>
  <a:clrScheme name='Custom'>
   <a:dk1><a:sysClr val='windowText' lastClr='000000'/></a:dk1>
   <a:lt1><a:sysClr val='window' lastClr='FFFFFF'/></a:lt1>
  </a:clrScheme>
  <a:fontScheme name='Custom'>
   <a:majorFont><a:latin typeface='Arial'/></a:majorFont>
   <a:minorFont><a:latin typeface='Arial'/></a:minorFont>
  </a:fontScheme>
  <a:fmtScheme name='Custom'>
   <a:fillStyleLst><a:solidFill><a:srgbClr val='FFFFFF'/></a:solidFill></a:fillStyleLst>
   <a:lnStyleLst><a:ln w='9525'><a:solidFill><a:srgbClr val='000000'/></a:solidFill></a:ln></a:lnStyleLst>
   <a:effectStyleLst><a:effectStyle/></a:effectStyleLst>
  </a:fmtScheme>
 </a:themeElements>
</a:theme>"""

def create_pptx(title: str, content: str) -> io.BytesIO:
    buffer = io.BytesIO()
    slide_xml = SLIDE1_XML_TEMPLATE.format(title=title, content=content)
    with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as pptx:
        pptx.writestr('[Content_Types].xml', CONTENT_TYPES)
        pptx.writestr('_rels/.rels', RELS)
        pptx.writestr('ppt/presentation.xml', PRESENTATION_XML)
        pptx.writestr('ppt/_rels/presentation.xml.rels', PRESENTATION_RELS)
        pptx.writestr('ppt/slides/slide1.xml', slide_xml)
        pptx.writestr('ppt/slides/_rels/slide1.xml.rels', SLIDE1_RELS)
        pptx.writestr('ppt/slideLayouts/slideLayout1.xml', SLIDE_LAYOUT_XML)
        pptx.writestr('ppt/slideLayouts/_rels/slideLayout1.xml.rels', SLIDE_LAYOUT_RELS)
        pptx.writestr('ppt/slideMasters/slideMaster1.xml', SLIDE_MASTER_XML)
        pptx.writestr('ppt/slideMasters/_rels/slideMaster1.xml.rels', SLIDE_MASTER_RELS)
        pptx.writestr('ppt/theme/theme1.xml', THEME1_XML)
    buffer.seek(0)
    return buffer

@app.post('/generate-ppt')
async def generate_ppt(title: str = Form(...), content: str = Form(...)):
    buf = create_pptx(title, content)
    headers = {'Content-Disposition': 'attachment; filename=presentation.pptx'}
    return StreamingResponse(buf, media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation', headers=headers)

if __name__ == '__main__':
    import uvicorn
    uvicorn.run(app, host='0.0.0.0', port=8000)
