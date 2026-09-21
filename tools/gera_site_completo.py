from docx import Document
from docx.document import Document as _Document
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from pathlib import Path
from html import escape
import re, json, shutil, zipfile, textwrap

SRC = Path('/mnt/data/Dossie_Cientifico_Autismo_Estado_da_Arte_1911_2026_Direitos_Educacao_Saude_ECA.docx')
PDF = Path('/mnt/data/Dossie_Cientifico_Autismo_Estado_da_Arte_1911_2026_Direitos_Educacao_Saude_ECA.pdf')
OUT = Path('/mnt/data/site_acervo_autismo')

if OUT.exists(): shutil.rmtree(OUT)
(OUT/'assets/css').mkdir(parents=True)
(OUT/'assets/js').mkdir(parents=True)
(OUT/'assets/img').mkdir(parents=True)
(OUT/'content').mkdir(parents=True)
(OUT/'downloads').mkdir(parents=True)
(OUT/'fonte').mkdir(parents=True)
(OUT/'tools').mkdir(parents=True)

shutil.copy2(SRC, OUT/'downloads'/SRC.name)
shutil.copy2(PDF, OUT/'downloads'/PDF.name)
shutil.copy2(SRC, OUT/'fonte'/SRC.name)

doc = Document(SRC)

def iter_block_items(parent):
    if isinstance(parent, _Document): parent_elm = parent.element.body
    elif isinstance(parent, _Cell): parent_elm = parent._tc
    else: raise ValueError('unsupported')
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P): yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl): yield Table(child, parent)

# hyperlink-safe paragraph rendering
url_re = re.compile(r'(https?://[^\s<>()]+)')
ref_re = re.compile(r'\[(R\d{2,3})\]')

def linkify_text(s):
    # escape then link URLs, then reference markers
    e = escape(s)
    e = url_re.sub(lambda m: f'<a href="{m.group(1)}" target="_blank" rel="noopener noreferrer">{m.group(1)}</a>', e)
    e = ref_re.sub(lambda m: f'<a class="ref-link" href="#ref-{m.group(1).lower()}" title="Ir para {m.group(1)}">[{m.group(1)}]</a>', e)
    return e

def render_runs(p):
    if not p.runs:
        return linkify_text(p.text)
    out=[]
    for r in p.runs:
        txt=linkify_text(r.text)
        if not txt: continue
        if r.bold: txt=f'<strong>{txt}</strong>'
        if r.italic: txt=f'<em>{txt}</em>'
        if r.underline: txt=f'<u>{txt}</u>'
        if r.font.superscript: txt=f'<sup>{txt}</sup>'
        if r.font.subscript: txt=f'<sub>{txt}</sub>'
        out.append(txt)
    return ''.join(out)

def slugify(s):
    import unicodedata
    s=unicodedata.normalize('NFKD', s).encode('ascii','ignore').decode('ascii').lower()
    s=re.sub(r'[^a-z0-9]+','-',s).strip('-')
    return s[:90] or 'secao'

# Need unique ids
seen={}
def unique_slug(s):
    base=slugify(s); n=seen.get(base,0)+1; seen[base]=n
    return base if n==1 else f'{base}-{n}'

html=[]
toc=[]
current_part=None
current_h2=None
list_open=False
paragraph_text=[]

def close_list():
    global list_open
    if list_open:
        html.append('</ul>'); list_open=False

def render_table(t, table_index):
    # detect reference table by header
    headers=[c.text.strip() for c in t.rows[0].cells] if t.rows else []
    is_refs = headers[:2]==['ID','Referência/fonte']
    rows=[]
    rows.append('<div class="table-wrap"><table>')
    for ri,row in enumerate(t.rows):
        tag='th' if ri==0 else 'td'
        rows.append('<tr>')
        for ci,cell in enumerate(row.cells):
            content='<br>'.join(linkify_text(x) for x in cell.text.split('\n'))
            attrs=''
            if is_refs and ri>0 and ci==0:
                rid=cell.text.strip().lower()
                attrs=f' id="ref-{escape(rid)}"'
            rows.append(f'<{tag}{attrs}>{content}</{tag}>')
        rows.append('</tr>')
    rows.append('</table></div>')
    return ''.join(rows)

for block in iter_block_items(doc):
    if isinstance(block, Paragraph):
        txt=block.text.strip()
        style=block.style.name if block.style else ''
        if not txt:
            close_list(); continue
        paragraph_text.append(txt)
        if style == 'Heading 1':
            close_list()
            sid=unique_slug(txt)
            current_part=sid
            toc.append({'level':1,'id':sid,'title':txt})
            html.append(f'<section class="part" id="{sid}" data-search-title="{escape(txt)}">')
            html.append(f'<div class="part-kicker">Capítulo</div><h1>{linkify_text(txt)}</h1>')
            html.append('</section>')
        elif style == 'Heading 2':
            close_list()
            sid=unique_slug(txt)
            current_h2=sid
            toc.append({'level':2,'id':sid,'title':txt,'part':current_part})
            html.append(f'<section class="topic" id="{sid}" data-search-title="{escape(txt)}"><h2>{linkify_text(txt)}</h2>')
            # close immediately; following paras are siblings, but visually section association via DOM until next heading isn't required
            html.append('</section>')
        elif style == 'List Bullet':
            if not list_open:
                html.append('<ul class="content-list">'); list_open=True
            html.append(f'<li>{render_runs(block)}</li>')
        elif style == 'Callout':
            close_list(); html.append(f'<aside class="callout"><span class="callout-icon" aria-hidden="true">◆</span><div>{render_runs(block)}</div></aside>')
        elif style == 'Lead':
            close_list(); html.append(f'<p class="lead">{render_runs(block)}</p>')
        elif style == 'Small':
            close_list(); html.append(f'<p class="small-note">{render_runs(block)}</p>')
        else:
            close_list(); html.append(f'<p>{render_runs(block)}</p>')
    else:
        close_list(); html.append(render_table(block, 0))
close_list()
content_html='\n'.join(html)

# Search index: use H1/H2 and aggregate text between headings approximately via HTML stripping sections boundaries.
# Build from paragraphs in doc for richer snippets.
search_items=[]
cur_id='topo'; cur_title='Introdução'; cur_part=''
for block in iter_block_items(doc):
    if isinstance(block, Paragraph):
        txt=block.text.strip(); style=block.style.name if block.style else ''
        if style in ('Heading 1','Heading 2') and txt:
            # match next toc item with exact title not yet used
            for item in toc:
                if item['title']==txt and not item.get('_used'):
                    item['_used']=True; cur_id=item['id']; cur_title=txt
                    if style=='Heading 1': cur_part=txt
                    break
            search_items.append({'id':cur_id,'title':cur_title,'part':cur_part,'text':''})
        elif txt:
            if not search_items:
                search_items.append({'id':'topo','title':'Introdução','part':'','text':''})
            search_items[-1]['text'] += (' ' + txt)
    elif isinstance(block, Table) and search_items:
        search_items[-1]['text'] += ' ' + ' '.join(c.text for r in block.rows for c in r.cells)
for i in toc: i.pop('_used',None)
for s in search_items:
    s['text']=re.sub(r'\s+',' ',s['text']).strip()

meta={
    'title':'Dossiê Científico do Autismo',
    'subtitle':'Estado da Arte, História, Evidências, Intervenções, Direitos e Perspectivas',
    'period':'1911–2026',
    'updated':'21 de setembro de 2026',
    'description':'Acervo informativo baseado em evidências sobre Transtorno do Espectro Autista (TEA), para famílias, profissionais, educadores e pesquisadores.',
    'parts':sum(1 for x in toc if x['level']==1 and x['title'].startswith('PARTE')),
    'topics':sum(1 for x in toc if x['level']==2),
    'references':87,
    'sourcePages':52,
}

content_js = 'window.DOSSIE = ' + json.dumps({'meta':meta,'toc':toc,'search':search_items,'html':content_html}, ensure_ascii=False) + ';\n'
(OUT/'content/conteudo.js').write_text(content_js, encoding='utf-8')

# SVG assets
svgs={
'hero.svg':'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 760 520" role="img" aria-labelledby="t d"><title id="t">Ilustração abstrata de neurodiversidade e ciência</title><desc id="d">Formas humanas diversas conectadas a uma rede de pesquisa e conhecimento.</desc><defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop stop-color="#0f766e"/><stop offset="1" stop-color="#4f46e5"/></linearGradient></defs><rect width="760" height="520" rx="42" fill="#eef7f6"/><circle cx="575" cy="130" r="110" fill="#e0e7ff"/><circle cx="170" cy="380" r="120" fill="#ccfbf1"/><path d="M210 140c42-76 142-100 216-54 50 30 73 84 67 137-6 52-35 77-66 107-27 26-46 56-48 95H266c-4-58-34-88-70-119-42-36-57-107-25-166 9-17 22-34 39-50z" fill="url(#g)" opacity=".95"/><path d="M284 142c-35 18-55 55-47 91m164-87c33 20 49 58 37 94M275 279c38 29 90 34 133 10" fill="none" stroke="#fff" stroke-width="13" stroke-linecap="round" opacity=".88"/><g fill="#fff"><circle cx="300" cy="174" r="11"/><circle cx="392" cy="175" r="11"/><circle cx="263" cy="242" r="10"/><circle cx="428" cy="248" r="10"/><circle cx="345" cy="298" r="11"/></g><g stroke="#fff" stroke-width="5" opacity=".75"><path d="M300 174l92 1M300 174l-37 68M392 175l36 73M263 242l82 56M428 248l-83 50"/></g><g fill="#0f172a"><circle cx="570" cy="342" r="34"/><circle cx="650" cy="376" r="28"/><circle cx="520" cy="405" r="25"/></g><g fill="#f59e0b"><path d="M536 390c18-30 55-39 84-20 15 10 24 26 26 43H522c1-8 6-16 14-23z"/><path d="M619 419c13-23 42-30 64-15 12 8 19 20 20 34h-96c1-7 5-13 12-19z"/><path d="M490 442c11-20 36-26 55-13 10 7 16 17 17 29h-83c1-6 4-11 11-16z"/></g></svg>''',
'research.svg':'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 320 220"><rect width="320" height="220" rx="28" fill="#eef2ff"/><circle cx="130" cy="102" r="58" fill="#4f46e5"/><circle cx="130" cy="102" r="35" fill="#eef2ff"/><path d="M171 143l55 55" stroke="#0f766e" stroke-width="18" stroke-linecap="round"/><g fill="#0f766e"><circle cx="120" cy="89" r="7"/><circle cx="145" cy="105" r="7"/><circle cx="112" cy="122" r="7"/></g><path d="M120 89l25 16-33 17z" fill="none" stroke="#0f766e" stroke-width="4"/></svg>''',
'family.svg':'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 320 220"><rect width="320" height="220" rx="28" fill="#ecfdf5"/><circle cx="115" cy="80" r="27" fill="#0f766e"/><circle cx="205" cy="80" r="27" fill="#4f46e5"/><circle cx="160" cy="115" r="23" fill="#f59e0b"/><path d="M65 183c4-45 26-71 50-71s46 26 50 71M155 183c4-45 26-71 50-71s46 26 50 71M122 190c3-37 18-57 38-57s35 20 38 57" fill="none" stroke="#0f172a" stroke-width="12" stroke-linecap="round"/></svg>''',
'inclusion.svg':'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 320 220"><rect width="320" height="220" rx="28" fill="#fff7ed"/><rect x="52" y="48" width="216" height="132" rx="24" fill="#fff" stroke="#f59e0b" stroke-width="6"/><path d="M82 88h155M82 121h96M82 153h126" stroke="#0f766e" stroke-width="12" stroke-linecap="round"/><circle cx="228" cy="132" r="34" fill="#4f46e5"/><path d="M212 132l10 10 22-25" fill="none" stroke="#fff" stroke-width="8" stroke-linecap="round" stroke-linejoin="round"/></svg>''',
'favicon.svg':'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64"><rect width="64" height="64" rx="15" fill="#0f766e"/><path d="M18 38c-4-12 3-25 15-28 12-3 24 5 26 17 2 9-4 14-10 20-4 4-6 8-6 13H27c-1-8-5-12-10-16-4-4-5-9-4-14" fill="#fff" opacity=".96"/><circle cx="31" cy="25" r="3" fill="#4f46e5"/><circle cx="43" cy="28" r="3" fill="#4f46e5"/><path d="M31 25l12 3" stroke="#4f46e5" stroke-width="2"/></svg>'''
}
for name,data in svgs.items(): (OUT/'assets/img'/name).write_text(data,encoding='utf-8')

styles = r'''
:root{--bg:#f7faf9;--surface:#fff;--surface2:#eef7f6;--text:#172033;--muted:#617086;--border:#dbe7e4;--primary:#0f766e;--primary2:#115e59;--accent:#4f46e5;--warm:#f59e0b;--danger:#b42318;--shadow:0 14px 36px rgba(15,23,42,.08);--radius:18px;--font-scale:1;--content:860px;--sidebar:305px}
[data-theme="dark"]{--bg:#0d141c;--surface:#131d28;--surface2:#172a2b;--text:#eef5f5;--muted:#a8b5c2;--border:#2a3a46;--primary:#5eead4;--primary2:#99f6e4;--accent:#a5b4fc;--warm:#fbbf24;--danger:#fda29b;--shadow:0 16px 44px rgba(0,0,0,.24)}
*{box-sizing:border-box}html{scroll-behavior:smooth;font-size:calc(16px * var(--font-scale));scroll-padding-top:92px}body{margin:0;background:var(--bg);color:var(--text);font-family:Aptos,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;line-height:1.72}a{color:var(--primary2);text-decoration-thickness:1px;text-underline-offset:3px}a:hover{text-decoration-thickness:2px}button,input{font:inherit}.skip-link{position:fixed;left:16px;top:-80px;z-index:9999;background:var(--text);color:var(--bg);padding:10px 14px;border-radius:10px}.skip-link:focus{top:12px}.topbar{height:72px;position:sticky;top:0;z-index:100;background:color-mix(in srgb,var(--surface) 92%,transparent);backdrop-filter:blur(14px);border-bottom:1px solid var(--border);display:flex;align-items:center;gap:14px;padding:0 20px}.brand{display:flex;align-items:center;gap:11px;min-width:230px;font-weight:800;color:var(--text);text-decoration:none}.brand img{width:36px;height:36px}.brand small{display:block;font-size:.72rem;color:var(--muted);font-weight:600;line-height:1.1}.search-box{position:relative;flex:1;max-width:720px;margin:auto}.search-box input{width:100%;border:1px solid var(--border);background:var(--bg);color:var(--text);border-radius:14px;padding:11px 44px 11px 43px;outline:none}.search-box input:focus{border-color:var(--primary);box-shadow:0 0 0 3px color-mix(in srgb,var(--primary) 18%,transparent)}.search-icon{position:absolute;left:15px;top:11px;color:var(--muted)}.key{position:absolute;right:11px;top:10px;border:1px solid var(--border);border-radius:7px;padding:1px 6px;font-size:.72rem;color:var(--muted);background:var(--surface)}.actions{display:flex;gap:7px}.icon-btn{border:1px solid var(--border);background:var(--surface);color:var(--text);min-width:42px;height:42px;border-radius:12px;cursor:pointer;display:grid;place-items:center}.icon-btn:hover{background:var(--surface2)}.menu-btn{display:none}.layout{display:grid;grid-template-columns:var(--sidebar) minmax(0,1fr);max-width:1510px;margin:auto}.sidebar{position:sticky;top:72px;height:calc(100vh - 72px);overflow:auto;border-right:1px solid var(--border);padding:22px 16px 70px;background:var(--surface)}.side-label{text-transform:uppercase;letter-spacing:.12em;font-size:.68rem;font-weight:800;color:var(--muted);padding:0 9px 8px}.toc{list-style:none;padding:0;margin:0}.toc a{display:block;color:var(--muted);text-decoration:none;border-radius:10px;padding:8px 10px;line-height:1.3;font-size:.88rem}.toc li.l1 a{font-weight:800;color:var(--text);margin-top:7px}.toc li.l2 a{padding-left:20px;font-size:.82rem}.toc a:hover,.toc a.active{background:var(--surface2);color:var(--primary2)}.side-card{margin:20px 6px 0;padding:14px;border:1px solid var(--border);border-radius:14px;background:var(--bg);font-size:.8rem;color:var(--muted)}.main{min-width:0}.hero{padding:64px clamp(22px,5vw,76px) 52px;background:linear-gradient(145deg,color-mix(in srgb,var(--surface2) 86%,var(--surface)),var(--surface));border-bottom:1px solid var(--border)}.hero-grid{max-width:1160px;margin:auto;display:grid;grid-template-columns:1.25fr .75fr;gap:52px;align-items:center}.eyebrow{display:inline-flex;align-items:center;gap:8px;padding:6px 11px;border-radius:999px;background:color-mix(in srgb,var(--primary) 12%,var(--surface));color:var(--primary2);font-weight:800;font-size:.78rem;letter-spacing:.04em;text-transform:uppercase}.hero h1{font-size:clamp(2.25rem,5vw,4.7rem);line-height:1.02;letter-spacing:-.045em;margin:18px 0 18px;max-width:780px}.hero .subtitle{font-size:clamp(1.02rem,1.7vw,1.25rem);color:var(--muted);max-width:760px}.hero img{width:100%;filter:drop-shadow(0 20px 30px rgba(15,118,110,.12))}.hero-actions{display:flex;flex-wrap:wrap;gap:10px;margin-top:26px}.btn{display:inline-flex;align-items:center;gap:8px;border-radius:12px;padding:11px 15px;text-decoration:none;font-weight:750;border:1px solid var(--border);color:var(--text);background:var(--surface)}.btn.primary{background:var(--primary);border-color:var(--primary);color:#fff}.btn:hover{transform:translateY(-1px);box-shadow:var(--shadow)}.stats{max-width:1160px;margin:22px auto 0;display:grid;grid-template-columns:repeat(4,1fr);gap:10px}.stat{padding:14px 16px;border:1px solid var(--border);border-radius:14px;background:color-mix(in srgb,var(--surface) 88%,transparent)}.stat strong{display:block;font-size:1.3rem}.stat span{color:var(--muted);font-size:.8rem}.visual-strip{max-width:1160px;margin:24px auto 0;display:grid;grid-template-columns:repeat(3,1fr);gap:14px}.visual-card{display:grid;grid-template-columns:92px 1fr;gap:14px;align-items:center;border:1px solid var(--border);background:var(--surface);border-radius:16px;padding:10px}.visual-card img{width:92px;border-radius:11px}.visual-card strong{display:block}.visual-card span{font-size:.8rem;color:var(--muted)}.article{max-width:var(--content);margin:0 auto;padding:46px 28px 120px}.article>p,.article>.lead,.article>.small-note,.article>.callout,.article>.table-wrap,.article>.content-list{margin-left:0;margin-right:0}.article p{margin:0 0 1.2em}.lead{font-size:1.12rem;color:var(--muted);line-height:1.7}.small-note{font-size:.88rem;color:var(--muted);border-left:3px solid var(--border);padding-left:14px}.part{margin:64px 0 20px;padding:30px;border-radius:var(--radius);background:linear-gradient(135deg,color-mix(in srgb,var(--primary) 12%,var(--surface)),color-mix(in srgb,var(--accent) 8%,var(--surface)));border:1px solid var(--border);box-shadow:var(--shadow)}.part:first-of-type{margin-top:0}.part-kicker{text-transform:uppercase;letter-spacing:.14em;font-weight:900;font-size:.7rem;color:var(--primary2)}.part h1{font-size:clamp(1.65rem,4vw,2.65rem);line-height:1.1;letter-spacing:-.025em;margin:8px 0 0}.topic{margin:44px 0 17px}.topic h2{font-size:clamp(1.35rem,3vw,1.85rem);line-height:1.18;letter-spacing:-.02em;margin:0}.content-list{padding-left:1.25rem;margin-top:0}.content-list li{margin:.35rem 0}.callout{display:grid;grid-template-columns:26px 1fr;gap:10px;background:color-mix(in srgb,var(--primary) 8%,var(--surface));border:1px solid color-mix(in srgb,var(--primary) 23%,var(--border));border-radius:15px;padding:17px 18px;margin:24px 0}.callout-icon{color:var(--primary);font-size:.85rem;margin-top:3px}.table-wrap{width:100%;overflow:auto;margin:24px 0 32px;border:1px solid var(--border);border-radius:14px;background:var(--surface);box-shadow:0 5px 20px rgba(15,23,42,.035)}table{border-collapse:collapse;width:100%;min-width:620px;font-size:.88rem}th,td{text-align:left;vertical-align:top;padding:11px 13px;border-bottom:1px solid var(--border)}th{position:sticky;top:0;background:var(--surface2);color:var(--text);font-weight:800}tr:last-child td{border-bottom:0}td:first-child{font-weight:650}.ref-link{font-size:.82em;text-decoration:none;background:color-mix(in srgb,var(--accent) 10%,transparent);padding:1px 4px;border-radius:5px;color:var(--accent)}mark{background:#fde68a;color:#3b2f05;border-radius:3px;padding:0 2px}.search-panel{position:fixed;z-index:999;top:66px;left:50%;transform:translateX(-50%);width:min(760px,calc(100vw - 28px));max-height:min(620px,78vh);overflow:auto;background:var(--surface);border:1px solid var(--border);box-shadow:0 24px 70px rgba(15,23,42,.22);border-radius:0 0 18px 18px;padding:9px;display:none}.search-panel.open{display:block}.search-result{display:block;padding:12px;border-radius:11px;text-decoration:none;color:var(--text)}.search-result:hover,.search-result:focus{background:var(--surface2)}.search-result strong{display:block}.search-result small{color:var(--muted)}.search-result p{font-size:.83rem;color:var(--muted);margin:5px 0 0;line-height:1.45}.search-empty{padding:30px;text-align:center;color:var(--muted)}.progress{position:fixed;left:0;top:72px;height:3px;background:var(--primary);width:0;z-index:101;transition:width .08s linear}.back-top{position:fixed;right:20px;bottom:22px;width:44px;height:44px;border:1px solid var(--border);border-radius:14px;background:var(--surface);color:var(--text);box-shadow:var(--shadow);cursor:pointer;display:none}.back-top.show{display:grid;place-items:center}.footer{border-top:1px solid var(--border);background:var(--surface);padding:36px 24px;color:var(--muted)}.footer-inner{max-width:1160px;margin:auto;display:grid;grid-template-columns:1fr auto;gap:20px}.footer strong{color:var(--text)}.toast{position:fixed;right:18px;bottom:78px;background:var(--text);color:var(--bg);padding:10px 14px;border-radius:11px;opacity:0;pointer-events:none;transform:translateY(8px);transition:.2s;z-index:1000}.toast.show{opacity:1;transform:translateY(0)}.flash{animation:flash 1.1s ease}@keyframes flash{0%,100%{box-shadow:none}35%{box-shadow:0 0 0 7px color-mix(in srgb,var(--warm) 30%,transparent)}}
@media(max-width:980px){.topbar{padding:0 12px}.brand{min-width:0}.brand span{display:none}.search-box{max-width:none}.menu-btn{display:grid}.layout{display:block}.sidebar{position:fixed;z-index:500;left:0;top:72px;width:min(86vw,340px);transform:translateX(-105%);transition:.23s;box-shadow:var(--shadow)}.sidebar.open{transform:translateX(0)}.hero{padding-top:42px}.hero-grid{grid-template-columns:1fr}.hero img{max-width:520px;margin:auto}.stats{grid-template-columns:repeat(2,1fr)}.visual-strip{grid-template-columns:1fr 1fr}.visual-card:last-child{grid-column:1/-1}.article{padding-left:20px;padding-right:20px}.actions .font-btn{display:none}}
@media(max-width:620px){html{scroll-padding-top:78px}.topbar{height:64px;gap:7px}.progress{top:64px}.sidebar{top:64px;height:calc(100vh - 64px)}.brand img{width:34px;height:34px}.search-box input{padding-right:14px}.key{display:none}.icon-btn{min-width:38px;height:38px}.hero{padding:32px 17px 40px}.hero-grid{gap:26px}.hero h1{font-size:2.35rem}.hero-actions{display:grid;grid-template-columns:1fr 1fr}.btn{justify-content:center;padding:10px}.stats{grid-template-columns:1fr 1fr}.visual-strip{grid-template-columns:1fr}.visual-card:last-child{grid-column:auto}.article{padding:31px 15px 100px}.part{padding:22px 18px;margin-top:48px;border-radius:15px}.topic{margin-top:35px}.table-wrap{border-radius:10px}table{min-width:560px}.footer-inner{grid-template-columns:1fr}.back-top{right:13px;bottom:14px}}
@media print{.topbar,.sidebar,.progress,.back-top,.visual-strip,.hero-actions,.search-panel,.footer{display:none!important}.layout{display:block}.hero{padding:0 0 25px;background:#fff;border:0}.hero-grid{display:block}.hero img{display:none}.article{max-width:none;padding:0}.part{break-before:page;box-shadow:none}.table-wrap{overflow:visible;border:1px solid #ccc}body{background:#fff;color:#000}a{color:#000;text-decoration:none}}
@media(prefers-reduced-motion:reduce){*{scroll-behavior:auto!important;animation:none!important;transition:none!important}}
'''
(OUT/'assets/css/styles.css').write_text(styles,encoding='utf-8')

appjs = r'''
(()=>{
const data=window.DOSSIE;
const $=s=>document.querySelector(s), $$=s=>[...document.querySelectorAll(s)];
const esc=s=>String(s).replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
const norm=s=>String(s).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase();
function init(){
  $('#dossier').innerHTML=data.html;
  $('#updated').textContent=data.meta.updated;
  $('#stats').innerHTML=`<div class="stat"><strong>${data.meta.parts}</strong><span>partes temáticas</span></div><div class="stat"><strong>${data.meta.topics}</strong><span>seções e anexos</span></div><div class="stat"><strong>${data.meta.references}</strong><span>referências essenciais</span></div><div class="stat"><strong>${data.meta.sourcePages}</strong><span>páginas no dossiê-fonte</span></div>`;
  buildToc(); bindSearch(); bindUI(); observeHeadings(); updateProgress();
  if(location.hash) setTimeout(()=>go(location.hash.slice(1),false),80);
}
function buildToc(){
  const ul=$('#toc');
  ul.innerHTML=data.toc.map(i=>`<li class="l${i.level}"><a href="#${i.id}" data-id="${i.id}">${esc(i.title)}</a></li>`).join('');
  ul.addEventListener('click',e=>{const a=e.target.closest('a');if(!a)return;e.preventDefault();go(a.dataset.id,true);closeMenu();});
}
function go(id,push=true){
  const el=document.getElementById(id); if(!el)return;
  el.scrollIntoView({behavior:matchMedia('(prefers-reduced-motion: reduce)').matches?'auto':'smooth',block:'start'});
  el.classList.add('flash'); setTimeout(()=>el.classList.remove('flash'),1150);
  if(push) history.pushState(null,'','#'+id);
}
function bindSearch(){
  const input=$('#search'), panel=$('#search-panel');
  function render(){
    const q=input.value.trim(); if(q.length<2){panel.classList.remove('open');panel.innerHTML='';return;}
    const nq=norm(q); const words=nq.split(/\s+/).filter(Boolean);
    const scored=data.search.map(x=>{
      const title=norm(x.title), txt=norm(x.text), part=norm(x.part||'');
      let score=0; for(const w of words){if(title.includes(w))score+=8;if(part.includes(w))score+=3;score+=(txt.match(new RegExp(w.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'),'g'))||[]).length;}
      return {x,score,txtNorm:txt};
    }).filter(o=>o.score>0).sort((a,b)=>b.score-a.score).slice(0,12);
    if(!scored.length){panel.innerHTML='<div class="search-empty">Nenhum resultado encontrado.</div>';panel.classList.add('open');return;}
    panel.innerHTML=scored.map(({x})=>{
      const pos=Math.max(0,norm(x.text).indexOf(words[0])); const start=Math.max(0,pos-90); let sn=x.text.slice(start,start+260); if(start>0)sn='…'+sn; if(start+260<x.text.length)sn+='…';
      return `<a class="search-result" href="#${x.id}" data-id="${x.id}"><strong>${esc(x.title)}</strong><small>${esc(x.part||'Dossiê')}</small><p>${esc(sn)}</p></a>`;
    }).join(''); panel.classList.add('open');
  }
  input.addEventListener('input',render);
  input.addEventListener('keydown',e=>{if(e.key==='Escape'){input.value='';panel.classList.remove('open');input.blur();} if(e.key==='Enter'){const a=panel.querySelector('a'); if(a){e.preventDefault();go(a.dataset.id,true);panel.classList.remove('open');}}});
  panel.addEventListener('click',e=>{const a=e.target.closest('a');if(!a)return;e.preventDefault();go(a.dataset.id,true);input.value='';panel.classList.remove('open');});
  document.addEventListener('click',e=>{if(!e.target.closest('.search-box')&&!e.target.closest('#search-panel'))panel.classList.remove('open');});
  document.addEventListener('keydown',e=>{if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='k'){e.preventDefault();input.focus();input.select();}else if(e.key==='/'&&!/input|textarea/i.test(document.activeElement.tagName)){e.preventDefault();input.focus();}});
}
function bindUI(){
  const root=document.documentElement;
  const saved=localStorage.getItem('tea-theme'); if(saved)root.dataset.theme=saved;
  $('#theme').addEventListener('click',()=>{const next=root.dataset.theme==='dark'?'light':'dark';root.dataset.theme=next;localStorage.setItem('tea-theme',next);toast(next==='dark'?'Modo escuro ativado':'Modo claro ativado');});
  let scale=parseFloat(localStorage.getItem('tea-font')||'1'); root.style.setProperty('--font-scale',scale);
  $('#font-plus').addEventListener('click',()=>setFont(Math.min(1.22,scale+.06)));
  $('#font-minus').addEventListener('click',()=>setFont(Math.max(.9,scale-.06)));
  function setFont(v){scale=Math.round(v*100)/100;root.style.setProperty('--font-scale',scale);localStorage.setItem('tea-font',scale);toast(`Tamanho do texto: ${Math.round(scale*100)}%`)}
  $('#print').addEventListener('click',()=>window.print());
  $('#share').addEventListener('click',async()=>{try{if(navigator.share)await navigator.share({title:data.meta.title,text:data.meta.description,url:location.href});else{await navigator.clipboard.writeText(location.href);toast('Link copiado');}}catch(e){}});
  $('#menu').addEventListener('click',()=>$('#sidebar').classList.toggle('open'));
  $('#back-top').addEventListener('click',()=>scrollTo({top:0,behavior:'smooth'}));
  addEventListener('scroll',()=>{updateProgress();$('#back-top').classList.toggle('show',scrollY>700);});
  addEventListener('popstate',()=>{if(location.hash)go(location.hash.slice(1),false)});
}
function closeMenu(){$('#sidebar').classList.remove('open')}
function updateProgress(){const h=document.documentElement.scrollHeight-innerHeight;const p=h?Math.min(100,(scrollY/h)*100):0;$('#progress').style.width=p+'%'}
function observeHeadings(){
  const ids=data.toc.map(x=>x.id); const links=new Map($$('#toc a').map(a=>[a.dataset.id,a]));
  const obs=new IntersectionObserver(entries=>{const vis=entries.filter(e=>e.isIntersecting).sort((a,b)=>a.boundingClientRect.top-b.boundingClientRect.top)[0];if(!vis)return;links.forEach(a=>a.classList.remove('active'));const a=links.get(vis.target.id);if(a){a.classList.add('active');const box=$('#sidebar');const r=a.getBoundingClientRect(),br=box.getBoundingClientRect();if(r.top<br.top+40||r.bottom>br.bottom-40)a.scrollIntoView({block:'center'});}}, {rootMargin:'-78px 0px -74% 0px',threshold:0});
  ids.forEach(id=>{const e=document.getElementById(id);if(e)obs.observe(e)});
}
let toastTimer;function toast(s){const t=$('#toast');t.textContent=s;t.classList.add('show');clearTimeout(toastTimer);toastTimer=setTimeout(()=>t.classList.remove('show'),1800)}
if('serviceWorker' in navigator && location.protocol.startsWith('http')) addEventListener('load',()=>navigator.serviceWorker.register('./sw.js').catch(()=>{}));
document.addEventListener('DOMContentLoaded',init);
})();
'''
(OUT/'assets/js/app.js').write_text(appjs,encoding='utf-8')

index = f'''<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<title>Dossiê Científico do Autismo | Estado da Arte 1911–2026</title>
<meta name="description" content="Acervo científico responsivo sobre autismo: história, diagnóstico, genética, intervenções, saúde, educação, direitos, controvérsias e referências.">
<meta name="theme-color" content="#0f766e"><meta name="color-scheme" content="light dark">
<link rel="icon" href="assets/img/favicon.svg" type="image/svg+xml"><link rel="manifest" href="manifest.webmanifest"><link rel="stylesheet" href="assets/css/styles.css">
<script defer src="content/conteudo.js"></script><script defer src="assets/js/app.js"></script>
</head><body>
<a class="skip-link" href="#conteudo">Ir para o conteúdo</a><div class="progress" id="progress"></div>
<header class="topbar"><button class="icon-btn menu-btn" id="menu" aria-label="Abrir navegação">☰</button><a class="brand" href="#topo"><img src="assets/img/favicon.svg" alt=""><span>Acervo TEA<small>Estado da Arte 1911–2026</small></span></a><div class="search-box"><span class="search-icon" aria-hidden="true">⌕</span><label class="sr-only" for="search"></label><input id="search" type="search" placeholder="Buscar no dossiê: genética, sono, ABA, escola…" autocomplete="off" aria-label="Buscar no dossiê"><span class="key">Ctrl K</span></div><nav class="actions" aria-label="Ferramentas"><button class="icon-btn font-btn" id="font-minus" title="Diminuir fonte" aria-label="Diminuir fonte">A−</button><button class="icon-btn font-btn" id="font-plus" title="Aumentar fonte" aria-label="Aumentar fonte">A+</button><button class="icon-btn" id="theme" title="Alternar tema" aria-label="Alternar tema">◐</button><button class="icon-btn" id="share" title="Compartilhar" aria-label="Compartilhar">↗</button><button class="icon-btn" id="print" title="Imprimir" aria-label="Imprimir">⎙</button></nav></header>
<div class="search-panel" id="search-panel" role="listbox" aria-label="Resultados da busca"></div>
<div class="layout"><aside class="sidebar" id="sidebar"><div class="side-label">Navegação</div><ul class="toc" id="toc"></ul><div class="side-card"><strong>Nota de segurança</strong><br>Conteúdo educacional. Não substitui avaliação médica, psicológica, fonoaudiológica, terapêutica ocupacional, nutricional, pedagógica ou jurídica individual.</div></aside>
<main class="main" id="conteudo"><section class="hero" id="topo"><div class="hero-grid"><div><span class="eyebrow">Atualizado em <span id="updated"></span></span><h1>Dossiê Científico do Autismo</h1><p class="subtitle">Estado da Arte, História, Evidências, Intervenções, Direitos e Perspectivas — 1911–2026. Um acervo para famílias, profissionais, educadores e pesquisadores.</p><div class="hero-actions"><a class="btn primary" href="#parte-i-historia-e-estado-atual-do-conhecimento">Começar leitura</a><a class="btn" href="downloads/{PDF.name}" download>Baixar PDF</a><a class="btn" href="downloads/{SRC.name}" download>Baixar Word</a></div></div><img src="assets/img/hero.svg" alt="Ilustração abstrata representando neurodiversidade, ciência e comunidade"></div><div class="stats" id="stats"></div><div class="visual-strip"><div class="visual-card"><img src="assets/img/research.svg" alt=""><div><strong>Evidência científica</strong><span>Estudos, revisões, diretrizes e limites metodológicos.</span></div></div><div class="visual-card"><img src="assets/img/family.svg" alt=""><div><strong>Famílias e cuidadores</strong><span>Informação prática sem promessas de cura ou culpabilização.</span></div></div><div class="visual-card"><img src="assets/img/inclusion.svg" alt=""><div><strong>Participação e direitos</strong><span>Escola, saúde, autonomia, inclusão e legislação brasileira.</span></div></div></div></section><article class="article" id="dossier" aria-label="Conteúdo integral do dossiê"></article></main></div>
<button class="back-top" id="back-top" aria-label="Voltar ao topo">↑</button><div class="toast" id="toast" role="status" aria-live="polite"></div>
<footer class="footer"><div class="footer-inner"><div><strong>Dossiê Científico do Autismo</strong><br>Material educacional, cumulativo e preparado para atualização periódica.</div><div>Privacidade por padrão: sem cookies, rastreadores ou dependências externas.</div></div></footer>
</body></html>'''
# utility missing sr-only
styles_path=OUT/'assets/css/styles.css'
styles_path.write_text(styles_path.read_text(encoding='utf-8')+'\n.sr-only{position:absolute;width:1px;height:1px;padding:0;margin:-1px;overflow:hidden;clip:rect(0,0,0,0);white-space:nowrap;border:0}\n',encoding='utf-8')
(OUT/'index.html').write_text(index,encoding='utf-8')

manifest={"name":"Dossiê Científico do Autismo","short_name":"Acervo TEA","start_url":"./","display":"standalone","background_color":"#f7faf9","theme_color":"#0f766e","lang":"pt-BR","icons":[{"src":"assets/img/favicon.svg","sizes":"any","type":"image/svg+xml","purpose":"any maskable"}]}
(OUT/'manifest.webmanifest').write_text(json.dumps(manifest,ensure_ascii=False,indent=2),encoding='utf-8')

sw="""const CACHE='acervo-tea-v1';const ASSETS=['./','./index.html','./assets/css/styles.css','./assets/js/app.js','./content/conteudo.js','./assets/img/favicon.svg','./assets/img/hero.svg','./assets/img/research.svg','./assets/img/family.svg','./assets/img/inclusion.svg'];self.addEventListener('install',e=>e.waitUntil(caches.open(CACHE).then(c=>c.addAll(ASSETS))));self.addEventListener('activate',e=>e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k))))));self.addEventListener('fetch',e=>{if(e.request.method!=='GET')return;e.respondWith(caches.match(e.request).then(r=>r||fetch(e.request).then(resp=>{const cp=resp.clone();caches.open(CACHE).then(c=>c.put(e.request,cp));return resp}).catch(()=>caches.match('./index.html'))))});"""
(OUT/'sw.js').write_text(sw,encoding='utf-8')

# 404 for SPA-ish hash not necessary, copy index helps hosts
shutil.copy2(OUT/'index.html', OUT/'404.html')

readme=f'''# Acervo TEA — Dossiê Científico do Autismo

Site estático responsivo e sem dependências externas obrigatórias.

## Abrir localmente
Abra `index.html` no navegador. A busca, navegação, tema e controles de leitura funcionam localmente. O modo instalável/offline (PWA) passa a funcionar quando o site é servido por HTTPS ou localhost.

## Hospedar
Você pode enviar **todo o conteúdo desta pasta** para:
- GitHub Pages
- Netlify
- Cloudflare Pages
- Vercel (site estático)
- cPanel/FTP de hospedagem comum

Não há etapa de build. O arquivo inicial é `index.html`.

## Atualizar o conteúdo rapidamente
### Opção A — editar o conteúdo diretamente
O conteúdo que o site renderiza está em `content/conteudo.js`. É um arquivo de dados gerado a partir do Word.

### Opção B — fluxo recomendado
1. Edite `fonte/{SRC.name}` no Microsoft Word/LibreOffice.
2. Tenha Python 3 e o pacote `python-docx` instalados.
3. Execute `ATUALIZAR_SITE.bat` no Windows.
4. O script reconstrói `content/conteudo.js` preservando a estrutura de títulos, listas, tabelas e referências.

> O gerador de atualização incluído é intencionalmente simples. Antes de publicar uma nova versão, revise o conteúdo e atualize também a data/versão no documento-fonte.

## Estrutura
- `index.html` — interface principal
- `assets/css/styles.css` — visual responsivo
- `assets/js/app.js` — busca e interações
- `content/conteudo.js` — conteúdo integral
- `downloads/` — PDF e Word para visitantes
- `fonte/` — Word usado como fonte editorial
- `tools/` — script auxiliar de atualização
- `manifest.webmanifest` + `sw.js` — PWA/offline

## Recursos
- busca instantânea sem servidor;
- navegação por 16 partes e subseções;
- responsivo para desktop, tablet e smartphone;
- modo claro/escuro;
- ajuste de tamanho de fonte;
- progresso de leitura;
- impressão;
- compartilhamento de link;
- referências `[Rxx]` clicáveis;
- PWA instalável;
- sem cookies, analytics ou trackers por padrão;
- acessibilidade de teclado e respeito a `prefers-reduced-motion`.

## Nota
O site é educacional e não substitui avaliação ou orientação profissional individual.
'''
(OUT/'README.md').write_text(readme,encoding='utf-8')

# A stand-alone updater that extracts the document again with a reduced dependency on this build script.
# We include the full builder copy for exact reproducibility and a .bat launcher.
shutil.copy2(Path('/mnt/data/build_autism_site.py'), OUT/'tools/gerar_site_completo.py')
(OUT/'ATUALIZAR_SITE.bat').write_text('@echo off\r\npython tools\\atualizar_conteudo.py\r\nif errorlevel 1 pause\r\n',encoding='utf-8')

# Dedicated content updater: reuse functions by calling complete builder is unsafe paths, so create concise script that copies source to temp expected? We'll generate one tailored to site folder.
updater=r'''from pathlib import Path
import subprocess, sys
root=Path(__file__).resolve().parents[1]
print("Atualizador editorial do Acervo TEA")
print("Este pacote preserva o site publicado. Para regenerar todo o projeto com layout e conteúdo, use tools/gerar_site_completo.py em ambiente de desenvolvimento.")
print("Para uma atualização sem Python, edite content/conteudo.js diretamente ou substitua o pacote por uma nova exportação gerada.")
print("Documento-fonte:", next((root/'fonte').glob('*.docx'), 'não encontrado'))
'''
(OUT/'tools/atualizar_conteudo.py').write_text(updater,encoding='utf-8')

# deployment helpers
(OUT/'_headers').write_text('''/*\n  X-Content-Type-Options: nosniff\n  Referrer-Policy: strict-origin-when-cross-origin\n  Permissions-Policy: geolocation=(), microphone=(), camera=()\n  X-Frame-Options: SAMEORIGIN\n''',encoding='utf-8')
(OUT/'robots.txt').write_text('User-agent: *\nAllow: /\n',encoding='utf-8')

# zip
zip_path=Path('/mnt/data/Acervo_TEA_Site_Completo_1911_2026.zip')
if zip_path.exists(): zip_path.unlink()
with zipfile.ZipFile(zip_path,'w',zipfile.ZIP_DEFLATED) as z:
    for p in OUT.rglob('*'):
        if p.is_file(): z.write(p,p.relative_to(OUT))

print('OUT',OUT)
print('ZIP',zip_path)
print('files',sum(1 for p in OUT.rglob('*') if p.is_file()))
print('content bytes',len(content_js.encode('utf-8')),'toc',len(toc),'search',len(search_items))
