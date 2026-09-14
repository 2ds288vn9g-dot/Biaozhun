from __future__ import annotations
import copy, re
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree

W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'; NS={'w':W}; q=lambda x:'{%s}%s'%(W,x)
FIELD={'ICS':'ics','CCS':'ccs','StandNo':'standard_number','ReplaceT':'replaces_standard','StdName':'chinese_title','StdEnglishName':'english_title','YZBS':'international_alignment','LB':'draft_type','FileSelect':'patent_notice','FY':'pub_y','FM':'pub_m','FD':'pub_d','SY':'imp_y','SM':'imp_m','SD':'imp_d','FM2':'publisher_top','FM3':'publisher_bottom'}
YZBS_TEXT='(点击此处添加与国际标准一致性程度的标识)'
DEFAULT_PRIMARY_ASSOCIATION='西安市质量与标准化协会'
DEFAULT_ASSOCIATION='青海省标准化协会'

def xml(data):
    with ZipFile(BytesIO(data)) as z:return etree.fromstring(z.read('word/document.xml'))
def body(root):return root.find(q('body'))
def txt(e):return ''.join(e.xpath('.//w:t/text()',namespaces=NS))
def cover_end(b):
    # 标准文件存在“目次”“目录”或直接从“前言”开始等多种结构。
    # 依次定位正文起点；最后使用第一页分节符，仍不能确认才停止。
    for marker in ('目次','目录','前言'):
        for i,e in enumerate(b):
            # WPS 自动目录可能将“目 次”拆为多个 run、tab 或文本框节点。
            visible=re.sub(r'\s+','', ''.join(e.itertext()))
            if marker in visible: return i
    # WPS 自动目录可能只保存 TOC 域指令，不保存可检索的“目录”文字。
    for i,e in enumerate(b):
        if e.xpath('.//w:instrText[contains(., "TOC")]',namespaces=NS): return i
    # “段前分页”标记的是第二页首段，直接以该段作为正文起点。
    for i,e in enumerate(b):
        if e.xpath('.//w:pageBreakBefore',namespaces=NS): return i
    # 普通分页符位于封面末段中，保留该段后从下一节点开始。
    for i,e in enumerate(b):
        if e.xpath('.//w:br[@w:type="page"]|.//w:lastRenderedPageBreak',namespaces=NS): return i+1
    for i,e in enumerate(b):
        if e.xpath('.//w:sectPr',namespaces=NS): return i+1
    raise ValueError('未找到封面结束位置，为保护正文已停止生成。')
def dateparts(s):
    m=re.search(r'((?:20\d{2}|X{4}))\D+((?:\d{1,2}|X{2}))\D+((?:\d{1,2}|X{2}))',s or '',re.I)
    return m.groups() if m else ('','','')
def set_legacy(root, values):
    values=dict(values); values.update(dict(zip(('pub_y','pub_m','pub_d'),dateparts(values.get('publication_date',''))))); values.update(dict(zip(('imp_y','imp_m','imp_d'),dateparts(values.get('implementation_date','')))))
    # 右上角编号下方的模板窗体域同样显示标准编号。
    values['replaces_standard']=values.get('standard_number','')
    values['publisher_bottom']=(values.get('publisher_bottom') or DEFAULT_ASSOCIATION).strip()
    for ff in root.xpath('.//w:ffData',namespaces=NS):
        name=ff.find(q('name')); key=FIELD.get(name.get(q('val')) if name is not None else '')
        if not key:continue
        value=values.get(key,'')
        # 一致性标识保留可编辑窗体域，但不再显示模板提示语。
        if key=='international_alignment': value=values.get(key,'')
        default=ff.find('.//'+q('default'))
        if default is not None:default.set(q('val'),value)
        if key=='standard_number':
            calc=ff.find(q('calcOnExit'))
            if calc is not None:calc.set(q('val'),'1')
        begin=ff.xpath('ancestor::w:r[1]',namespaces=NS)
        if not begin:continue
        p=begin[0].getparent(); children=list(p); start=children.index(begin[0]); state='before'; result=[]
        # 一个日期段落有三个字段；必须从当前 ffData 所在 begin run 开始，不能遍历整段。
        for r in children[start+1:]:
            if r.xpath('.//w:fldChar[@w:fldCharType="separate"]',namespaces=NS):state='result';continue
            if r.xpath('.//w:fldChar[@w:fldCharType="end"]',namespaces=NS):break
            if state=='result':result += r.xpath('.//w:t',namespaces=NS)
        if result:
            result[0].text=value if value else '\u2002\u2002\u2002\u2002\u2002'
            for t in result[1:]:t.text=''

def preserve_ics_ccs_spacing(root):
    """将模板文档网格的行距显式落到 ICS/CCS 两个表格行上。
    合并到原文件后，原文件的文档网格会覆盖模板的隐式行距；显式行高可保持模板效果。"""
    grid=root.find('.//'+q('docGrid'))
    pitch=grid.get(q('linePitch')) if grid is not None else None
    pitch=pitch or '579'  # 模板标准封面默认约 29 磅一行
    for ff in root.xpath('.//w:ffData',namespaces=NS):
        name=ff.find(q('name'))
        if name is None or name.get(q('val')) not in {'ICS','CCS'}: continue
        tables=ff.xpath('ancestor::w:tbl[1]',namespaces=NS)
        if not tables: continue
        for tr in tables[0].findall(q('tr')):
            trpr=tr.find(q('trPr'))
            if trpr is None:
                trpr=etree.Element(q('trPr'))
                tr.insert(0,trpr)
            height=trpr.find(q('trHeight'))
            if height is None:
                height=etree.SubElement(trpr,q('trHeight'))
                height.set(q('val'),pitch)
                height.set(q('hRule'),'atLeast')

def replace_preface_fields(root, values):
    """仅覆盖前言指定段落；提出与归口始终独立为两个段落。"""
    def set_paragraph(p, content):
        texts=p.xpath('.//w:t',namespaces=NS)
        if not texts: return False
        texts[0].text=content
        for t in texts[1:]: t.text=''
        return True

    association='、'.join(x for x in ((values.get('publisher_top') or DEFAULT_PRIMARY_ASSOCIATION).strip(), (values.get('publisher_bottom') or DEFAULT_ASSOCIATION).strip()) if x)
    proposal_paragraphs={}
    for p in root.xpath('.//w:p',namespaces=NS):
        current=txt(p)
        m=re.match(r'^\s*本文件由\s*.*?(提出|归口)\s*[。.]?\s*$',current)
        if m and m.group(1) not in proposal_paragraphs:
            proposal_paragraphs[m.group(1)]=p
    # 以原有“归口”或“提出”段落的格式为样式来源；缺失时复制它并插入相邻位置。
    reference=proposal_paragraphs.get('提出') or proposal_paragraphs.get('归口')
    if reference is not None:
        if '提出' not in proposal_paragraphs:
            p=copy.deepcopy(reference); reference.addprevious(p); proposal_paragraphs['提出']=p
        if '归口' not in proposal_paragraphs:
            p=copy.deepcopy(proposal_paragraphs['提出']); proposal_paragraphs['提出'].addnext(p); proposal_paragraphs['归口']=p
        set_paragraph(proposal_paragraphs['提出'], '本文件由'+association+'提出。')
        set_paragraph(proposal_paragraphs['归口'], '本文件由'+association+'归口。')

    for key,label in (('drafting_organizations','本文件起草单位'),('drafting_authors','本文件主要起草人')):
        value=(values.get(key) or '').strip()
        if not value: continue
        for p in root.xpath('.//w:p',namespaces=NS):
            current=txt(p)
            m=re.match(r'^(\s*'+re.escape(label)+r'\s*[：:])\s*.*$',current)
            if not m: continue
            end='。' if current.rstrip().endswith('。') else ''
            set_paragraph(p, m.group(1)+value.rstrip('。')+end)
            break

def replace_header_with_standard_ref(header, standard, number_re):
    """将页眉编号换成 REF StandNo 联动域，保留原首文字块的格式。"""
    for p in header.xpath('.//w:p',namespaces=NS):
        texts=p.xpath('.//w:t',namespaces=NS)
        joined=''.join(t.text or '' for t in texts)
        if not texts or not number_re.search(joined): continue
        first_run=texts[0].getparent()
        rpr=first_run.find(q('rPr')) if first_run is not None else None
        ppr=p.find(q('pPr'))
        for node in list(p):
            if node is not ppr:p.remove(node)
        def run(kind, value=None):
            r=etree.Element(q('r'))
            if rpr is not None:r.append(copy.deepcopy(rpr))
            if kind=='char':
                ch=etree.SubElement(r,q('fldChar'));ch.set(q('fldCharType'),value)
            elif kind=='instr':
                ins=etree.SubElement(r,q('instrText'));ins.set('{http://www.w3.org/XML/1998/namespace}space','preserve');ins.text=' REF StandNo \\* MERGEFORMAT '
            else:etree.SubElement(r,q('t')).text=value
            p.append(r)
        run('char','begin');run('instr');run('char','separate');run('text',standard);run('char','end')
def detect(source):
    root=xml(source); alltxt='\n'.join(txt(x) for x in body(root)[:45]); cells=[txt(x).strip() for x in root.xpath('.//w:tc',namespaces=NS)]
    def one(p):
        m=re.search(p,alltxt,re.I);return m.group(1).strip() if m else ''
    out={'ics':one(r'ICS\s*[:：]?\s*([0-9.]+)'),'ccs':one(r'CCS\s*[:：]?\s*([A-Z][A-Z0-9 /.-]+)'),'standard_number':one(r'((?:T|Q)/[^\s]+(?:\s+[^\s—-]+){0,3}\s*[—-]\s*(?:\d{4}|X{2,}))'),'replaces_standard':'','international_alignment':'','draft_type':'（工作组讨论稿）','patent_notice':'在提交反馈意见时，请将您知道的相关专利连同支持性文件一并附上。','publication_date':'','implementation_date':'','publisher_top':DEFAULT_PRIMARY_ASSOCIATION,'publisher_bottom':DEFAULT_ASSOCIATION,'chinese_title':'','english_title':'','drafting_organizations':'','drafting_authors':''}
    for i,c in enumerate(cells[:-1]):
        if c.upper()=='ICS' and not out['ics']:out['ics']=cells[i+1]
        if c.upper()=='CCS' and not out['ccs']:
            m=re.search(r'([A-Z]\s*\d{2}(?:\s*/\s*\d{2})?)\s*$',cells[i+1])
            out['ccs']=m.group(1) if m else cells[i+1]
    m=re.search(r'([A-Z])\s*\d{2}\s*/\s*(\d{2})\s*$',out['ccs'])
    if m: out['ccs']=m.group(1)+' '+m.group(2)
    else:
        m=re.search(r'([A-Z]\s*\d{2}(?:\s*/\s*\d{2})?)\s*$',out['ccs'])
        if m: out['ccs']=m.group(1)
    # 优先按页面实际显示的“发布/实施”整行读取，避免内部窗体域默认值与显示值不一致。
    date_line=re.compile(r'((?:20\d{2}|X{4}))\s*[-—–年]\s*((?:\d{1,2}|X{2}))\s*[-—–月]\s*((?:\d{1,2}|X{2}))\s*(?:日)?\s*(发布|实施)',re.I)
    for p in root.xpath('.//w:p',namespaces=NS):
        shown=''.join(p.xpath('.//w:t/text()',namespaces=NS))
        for year,month,day,label in date_line.findall(shown):
            out['publication_date' if label=='发布' else 'implementation_date']=' - '.join((year,month,day))
    # 日期在源文件中由六个独立窗体域保存；优先读取它们，可完整保留 XX 占位日期。
    legacy={}
    for ff in root.xpath('.//w:ffData',namespaces=NS):
        name=ff.find(q('name'))
        default=ff.find('.//'+q('default'))
        if name is None: continue
        # 优先读取窗体域在文档中实际显示的结果文本；默认值仅作为备用。
        displayed=''
        begin=ff.xpath('ancestor::w:r[1]',namespaces=NS)
        if begin:
            p=begin[0].getparent(); children=list(p); start=children.index(begin[0]); state='before'; result=[]
            for r in children[start+1:]:
                if r.xpath('.//w:fldChar[@w:fldCharType="separate"]',namespaces=NS): state='result'; continue
                if r.xpath('.//w:fldChar[@w:fldCharType="end"]',namespaces=NS): break
                if state=='result': result += r.xpath('.//w:t/text()',namespaces=NS)
            displayed=''.join(result).strip()
        legacy[name.get(q('val'))]=displayed or (default.get(q('val'),'') if default is not None else '')
    for target, names in (('publication_date',('FY','FM','FD')),('implementation_date',('SY','SM','SD'))):
        parts=[legacy.get(name,'').strip() for name in names]
        if not out[target] and all(parts): out[target]=' - '.join(parts)
    lines=[x.strip() for x in alltxt.splitlines() if x.strip()]; n=next((i for i,x in enumerate(lines) if out['standard_number'] and out['standard_number'] in x),-1)
    for x in lines[n+1:n+9] if n>=0 else lines[:15]:
        if not out['chinese_title'] and re.search('[\u4e00-\u9fff]',x) and '团体标准' not in x and len(x)>5:out['chinese_title']=x
        if not out['english_title'] and len(re.sub('[^A-Za-z]','',x))>20:out['english_title']=x
        if not out['publisher_top'] and '团体标准' in x:out['publisher_top']=x.replace('团体标准','').strip()
    if not out['publisher_top']:
        for x in lines[:20]:
            if '团体标准' in x:
                out['publisher_top']=x.replace('团体标准','').strip();break
    if not out['publisher_top']:
        for x in lines[:50]:
            m=re.match(r'(.+?)\s*发布$',x)
            if m and not re.match(r'20\d{2}',m.group(1).strip()):
                out['publisher_top']=m.group(1).strip();break
    for p in root.xpath('.//w:p',namespaces=NS):
        line=txt(p).strip()
        m=re.match(r'^本文件起草单位\s*[：:]\s*(.+?)(?:。)?$',line)
        if m and not out['drafting_organizations']:out['drafting_organizations']=m.group(1).rstrip('。').strip()
        m=re.match(r'^本文件主要起草人\s*[：:]\s*(.+?)(?:。)?$',line)
        if m and not out['drafting_authors']:out['drafting_authors']=m.group(1).rstrip('。').strip()
    out['replaces_standard']=out['standard_number']
    out['publisher_bottom']=out.get('publisher_bottom') or DEFAULT_ASSOCIATION
    return out
def merge(template, source, values):
    tz=ZipFile(BytesIO(template)); sz=ZipFile(BytesIO(source)); tr=etree.fromstring(tz.read('word/document.xml')); sr=etree.fromstring(sz.read('word/document.xml')); set_legacy(tr,values); preserve_ics_ccs_spacing(tr); replace_preface_fields(sr,values)
    tb,sb=body(tr),body(sr); cover=[copy.deepcopy(x) for x in list(tb)[:cover_end(tb)]]
    # “目次”前的模板内容已经自带封面结束分节符。再插入一次会产生空白页，
    # 因此保留模板原有分节符，不额外创建新的分页段落。
    if not any(x.find('.//'+q('sectPr')) is not None for x in cover):
        raise RuntimeError('模板封面缺少页面分节信息，已停止生成。')
    # 封面模板的页眉/页脚关系 ID 属于模板文件，不能带入待修改文件；否则 WPS
    # 会出现“未找到关联章节样式”。正文页眉继续使用待修改文件自己的关系与格式。
    for sect in [s for e in cover for s in e.xpath('.//w:sectPr',namespaces=NS)]:
        for ref in list(sect.findall(q('headerReference')))+list(sect.findall(q('footerReference'))):sect.remove(ref)
    # 给模板封面样式换独立 ID，绝不覆盖原文件正文正在使用的样式。
    ts=etree.fromstring(tz.read('word/styles.xml')); ss=etree.fromstring(sz.read('word/styles.xml')); defs={x.get(q('styleId')):x for x in ts.findall(q('style'))}; used={x.get(q('val')) for e in cover for x in e.xpath('.//w:pStyle|.//w:rStyle|.//w:tblStyle',namespaces=NS)}; stack=list(used)
    while stack:
        st=defs.get(stack.pop())
        if st is not None:
            b=st.find(q('basedOn')); v=b.get(q('val')) if b is not None else None
            if v and v not in used:used.add(v);stack.append(v)
    mapping={old:'cover_'+str(i) for i,old in enumerate(sorted(used))}
    for e in cover:
        for x in e.xpath('.//w:pStyle|.//w:rStyle|.//w:tblStyle',namespaces=NS):
            if x.get(q('val')) in mapping:x.set(q('val'),mapping[x.get(q('val'))])
    # 样式本身也可能继续引用其他样式（basedOn/next/link，或表格样式内部
    # 的段落样式）。统一改写所有引用，避免源文件同名样式重新覆盖封面行距。
    ref_nodes = './/w:basedOn|.//w:next|.//w:link|.//w:pStyle|.//w:rStyle|.//w:tblStyle'
    for old,new in mapping.items():
        st=copy.deepcopy(defs[old]);st.set(q('styleId'),new)
        for ref in st.xpath(ref_nodes,namespaces=NS):
            val=ref.get(q('val'))
            if val in mapping: ref.set(q('val'),mapping[val])
        ss.append(st)
    suffix=list(sb)[cover_end(sb):]
    for x in list(sb)[:-1]:sb.remove(x)
    for x in reversed(cover+suffix):sb.insert(0,x)
    # second page onwards exact source nodes are preserved; assertion blocks unsafe output.
    now=list(sb)[len(cover):]
    if len(now)!=len(suffix) or any(etree.tostring(a,method='c14n')!=etree.tostring(b,method='c14n') for a,b in zip(now,suffix)):raise RuntimeError('正文校验失败，已停止生成。')
    out=BytesIO()
    standard=values.get('standard_number','').strip()
    number_re=re.compile(r'(?:T|Q)/[^\s]+(?:\s+[^\s—-]+){0,3}\s*[—-]\s*(?:\d{4}|X{2,})',re.I)
    with ZipFile(out,'w',ZIP_DEFLATED) as z:
        for i in sz.infolist():
            if i.filename=='word/document.xml': data=etree.tostring(sr,xml_declaration=True,encoding='UTF-8',standalone=True)
            elif i.filename=='word/styles.xml': data=etree.tostring(ss,xml_declaration=True,encoding='UTF-8',standalone=True)
            elif standard and re.fullmatch(r'word/header\d+\.xml',i.filename):
                header=etree.fromstring(sz.read(i.filename))
                # 页眉编号可能被拆成多个 w:t，或旧格式为 Q/…；按整个页眉段落
                # 定位后直接以新标准编号覆盖，保留原页眉第一文字块的格式和位置。
                replace_header_with_standard_ref(header,standard,number_re)
                data=etree.tostring(header,xml_declaration=True,encoding='UTF-8',standalone=True)
            elif i.filename=='word/settings.xml':
                settings=etree.fromstring(sz.read(i.filename))
                update=settings.find(q('updateFields'))
                if update is None:update=etree.SubElement(settings,q('updateFields'))
                update.set(q('val'),'true')
                data=etree.tostring(settings,xml_declaration=True,encoding='UTF-8',standalone=True)
            else: data=sz.read(i.filename)
            z.writestr(i,data)
    return out.getvalue()
