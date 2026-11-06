"""Native numeric citation fields and independent XML evidence/protection."""
from __future__ import annotations

import copy
import re

from docx.oxml import OxmlElement

from .citation_display import DISPLAY_VALUES, run_display, write_display
from .contracts import FormatError
from .namespaces import qn
from .properties import element, strip_text_prefix, xml_value

PREFIX = re.compile(r'^\s*(?:\[\s*\d+\s*\]|［\s*\d+\s*］|\d+[.)、])\s*')


def numeric_tokens(text):
    """One complete native label per REF, portable across Word and WPS."""
    group=r'\[([1-9]\d*)\]'
    if not re.fullmatch(group+r'(?:\s*[,，、;；]?\s*'+group+r')*',text):
        raise FormatError('Dynamic numeric citations require a complete label per reference, e.g. [1][2]; grouped ranges/author-date are unsupported','unsupported')
    tokens=list(re.finditer(group,text))
    numbers=[int(t[1]) for t in tokens]
    if any(n>100000 for n in numbers):raise ValueError('Invalid numeric reference')
    return numbers,tokens


def _marker(document, prefix):
    ids = [int(n.get(qn('w:id'))) for n in document.element.iter(qn('w:bookmarkStart'))
           if (n.get(qn('w:id')) or '').isdigit()]
    names = {n.get(qn('w:name')) for n in document.element.iter(qn('w:bookmarkStart'))}
    number = max(ids, default=0)+1
    while prefix+str(number) in names:
        number += 1
    return number, prefix+str(number)


def _bookmark(document, paragraph):
    existing = next((n for n in paragraph.iter(qn('w:bookmarkStart'))
                     if (n.get(qn('w:name')) or '').startswith('WFMRef_')), None)
    if existing is not None:
        return existing.get(qn('w:name'))
    ident, name = _marker(document,'WFMRef_')
    start=OxmlElement('w:bookmarkStart');start.set(qn('w:id'),str(ident));start.set(qn('w:name'),name)
    end=OxmlElement('w:bookmarkEnd');end.set(qn('w:id'),str(ident))
    paragraph.insert(1 if paragraph.pPr is not None else 0,start)
    paragraph.append(end)
    return name


def _number_references(document, references):
    """One native sequence; retain reference typography and paragraph style."""
    from .structures import _get_numbering_element, _numbering_max_id
    from .properties import PARAGRAPH_TAGS, read_paragraph, write_paragraph, write_effective_tabs
    from .numbering import effective_numbering

    retained=[]
    targets={document.paragraphs[r['paragraph_index']]._p for r in references}
    native_ids=set();marker_fonts=[]
    for ref in references:
        p=document.paragraphs[ref['paragraph_index']]._p
        native=effective_numbering(document,p)
        if native:
            native_ids.add(native[0])
            level=next((n for n in native[2].findall(qn('w:lvl')) if n.get(qn('w:ilvl'))==str(native[1])),None) if native[2] is not None else None
            override=next((n for n in native[3].findall(qn('w:lvlOverride')) if n.get(qn('w:ilvl'))==str(native[1])),None)
            if override is not None and override.find(qn('w:lvl')) is not None:level=override.find(qn('w:lvl'))
            marker_fonts.append(level.find(qn('w:rPr')) if level is not None else None)
        elif PREFIX.match(document.paragraphs[ref['paragraph_index']].text):
            runs=document.paragraphs[ref['paragraph_index']].runs
            marker_fonts.append(runs[0]._r.rPr if runs else None)
        retained.append({k:read_paragraph(document,p,k) for k in PARAGRAPH_TAGS
                         if k not in {'numbering_suppressed','outline_level','hanging_indent'}})
    for p in document.element.iter(qn('w:p')):
        if p in targets:continue
        native=effective_numbering(document,p)
        if native and native[0] in native_ids:
            raise FormatError('Reference numbering shares a sequence outside bibliography; explicit separation required','unsupported')
    if len({xml_value(pr) for pr in marker_fonts})>1:
        raise FormatError('Reference numbering has mixed marker typography; explicit normalization required','unsupported')

    root=_get_numbering_element(document,create=True)
    aid=_numbering_max_id(root,'abstractNum','abstractNumId')+1
    nid=max(0,_numbering_max_id(root,'num','numId'))+1
    abstract=OxmlElement('w:abstractNum');abstract.set(qn('w:abstractNumId'),str(aid))
    element(abstract,'multiLevelType').set(qn('w:val'),'singleLevel')
    level=element(abstract,'lvl');level.set(qn('w:ilvl'),'0')
    for tag,value in [('start','1'),('numFmt','decimal'),('lvlText','[%1]'),('suff','space')]:
        element(level,tag).set(qn('w:val'),value)
    if marker_fonts and marker_fonts[0] is not None:level.append(copy.deepcopy(marker_fonts[0]))
    root.append(abstract)
    num=OxmlElement('w:num');num.set(qn('w:numId'),str(nid))
    element(num,'abstractNumId').set(qn('w:val'),str(aid));root.append(num)
    for ref,values in zip(references,retained):
        paragraph=document.paragraphs[ref['paragraph_index']]
        prefix=PREFIX.match(paragraph.text)
        if prefix:
            strip_text_prefix(paragraph._p,prefix.end())
        np=element(paragraph._p.get_or_add_pPr(),'numPr')
        for child in list(np):np.remove(child)
        element(np,'ilvl').set(qn('w:val'),'0')
        element(np,'numId').set(qn('w:val'),str(nid))
        for key,value in values.items():
            if read_paragraph(document,paragraph._p,key)==value:continue
            pr=paragraph._p.get_or_add_pPr()
            if key=='tabs':write_effective_tabs(document,paragraph._p,pr,value)
            elif key=='first_line_indent' and value['value']<0:
                write_paragraph(pr,'hanging_indent',{**value,'value':-value['value']})
            else:write_paragraph(pr,key,value)


def _run(text=None, pr=None):
    node=OxmlElement('w:r')
    if pr is not None:node.append(copy.deepcopy(pr))
    if text:
        t=element(node,'t');t.set(qn('xml:space'),'preserve');t.text=text
    return node


def _field(name, number, pr):
    result=[]
    for tag,value in [('fldChar','begin'),('instrText',f' REF {name} \\n \\h \\* MERGEFORMAT '),('fldChar','separate'),('t',f'[{number}]'),('fldChar','end')]:
        run=_run(pr=pr);node=element(run,tag)
        if tag=='fldChar':node.set(qn('w:fldCharType'),value)
        else:node.text=value
        result.append(run)
    return result


def _insert(paragraph, offset, nodes):
    """Split ordinary text only. Validation rejects complex anchors beforehand."""
    cursor=0
    for run in paragraph.runs:
        length=len(run.text)
        if offset <= cursor+length:
            before,after=run.text[:offset-cursor],run.text[offset-cursor:]
            suffix=copy.deepcopy(run._r)
            for child in list(suffix):
                if child.tag!=qn('w:rPr'):suffix.remove(child)
            if after:
                text=element(suffix,'t');text.text=after;text.set(qn('xml:space'),'preserve')
            run.text=before
            anchor=run._r
            for node in nodes:
                anchor.addnext(node);anchor=node
            if after:anchor.addnext(suffix)
            return
        cursor+=length
    for node in nodes:paragraph._p.append(node)


def insert_citations(document, citations):
    from citation_workflow import extract_citation_data, paragraph_sha256, _citation_offset

    if not citations:return []
    candidates,references=extract_citation_data(document)
    valid={p['paragraph_index'] for p in candidates}
    ids={r['id']:r for r in references}
    validated=[];seen=set()
    for placement in citations['placements']:
        index=placement['paragraph_index']
        if index not in valid or index in seen:
            raise ValueError('Citation must select one unambiguous body paragraph')
        seen.add(index);paragraph=document.paragraphs[index]
        if placement['paragraph_sha256']!=paragraph_sha256(paragraph.text):
            raise ValueError('Citation paragraph fingerprint mismatch')
        targets=placement['reference_ids']
        if any(r not in ids for r in targets) or len(set(targets))!=len(targets):
            raise ValueError('Citation references are invalid')
        display=placement['citation_text'].strip()
        numbers,tokens=numeric_tokens(display)
        expected=[int(ids[r]['id'].split('-')[-1]) for r in targets]
        if numbers!=expected:
            raise ValueError('Citation display numbers do not match reference_ids')
        if display in paragraph.text:raise ValueError('Citation text already exists')
        # Reject hyperlinks, fields, drawings, tabs, revisions and other complex runs.
        if any(c.tag not in {qn('w:pPr'),qn('w:r')} for c in paragraph._p) or any(
            c.tag not in {qn('w:rPr'),qn('w:t')} for r in paragraph.runs for c in r._r
        ):
            raise FormatError('Complex citation anchor is unsupported','unsupported')
        spacing=placement.get('spacing','none')
        if spacing not in {'none','space-before','space-after','spaces-around'}:
            raise ValueError('Invalid citation spacing')
        if placement.get('display', 'inline') not in tuple(DISPLAY_VALUES):
            raise ValueError('Invalid citation display')
        offset=_citation_offset(paragraph.text,placement)
        validated.append((paragraph,placement,display,tokens,offset,spacing))
    if not validated:return []
    _number_references(document,references)
    bookmarks={r['id']:_bookmark(document,document.paragraphs[r['paragraph_index']]._p) for r in references}
    for paragraph,placement,display,tokens,offset,spacing in validated:
        original_pr = paragraph.runs[0]._r.rPr if paragraph.runs else None
        pr = copy.deepcopy(original_pr) if original_pr is not None else OxmlElement('w:rPr')
        write_display(pr, placement.get('display', 'inline'))
        ident,name=_marker(document,f'WFMCite_{placement["paragraph_index"]}_')
        start=OxmlElement('w:bookmarkStart');start.set(qn('w:id'),str(ident));start.set(qn('w:name'),name)
        end=OxmlElement('w:bookmarkEnd');end.set(qn('w:id'),str(ident))
        nodes=[start]
        if spacing in {'space-before','spaces-around'}:nodes.append(_run(' ',original_pr))
        cursor=0
        for token in tokens:
            if token.start() > cursor:
                nodes.append(_run(display[cursor:token.start()],pr))
            number=int(token[1]);name=bookmarks['ref-'+str(number)]
            nodes.extend(_field(name,number,pr));cursor=token.end()
        if cursor < len(display):
            nodes.append(_run(display[cursor:],pr))
        if spacing in {'space-after','spaces-around'}:nodes.append(_run(' ',original_pr))
        nodes.append(end);_insert(paragraph,offset,nodes)
    return [f'Inserted {len(validated)} AI-approved citation placement(s) as dynamic REF fields with native references']


def read_state(document):
    """Independent reader of actual fields, targets and native numbering."""
    from .inspect import paragraph_text
    from .numbering import effective_numbering, sequence_number

    from .inspect import DocumentIndex
    index=DocumentIndex(document)
    ref_nodes=index.reference_nodes()
    kept=list(document.paragraphs)
    bookmarks={}
    for i,p in enumerate(kept):
        for start in p._p.iter(qn('w:bookmarkStart')):
            name=start.get(qn('w:name'))
            ends=[n for n in p._p.iter(qn('w:bookmarkEnd')) if n.get(qn('w:id'))==start.get(qn('w:id'))]
            bookmarks.setdefault(name,[]).append({'paragraph':i,'paired':len(ends)==1})
    paragraphs=[];counters={};reference_rows=[]
    for i,p in enumerate(kept):
        if p._p in ref_nodes and p.text.strip():reference_rows.append(i)
        stack=[];fields=[];offset=0
        for node in p._p.iter():
            if node.tag==qn('w:fldChar'):
                kind=node.get(qn('w:fldCharType'))
                if kind=='begin':
                    stack.append({'code':'','text':'','result':False,'start':None,'displays':[],
                                  'controls':set()})
                elif kind=='separate' and stack:
                    stack[-1].update(result=True, start=offset)
                if stack and kind in {'begin', 'separate', 'end'}:
                    stack[-1]['controls'].add(run_display(document, p._p, node.getparent()))
                if kind=='end' and stack:
                    field = stack.pop()
                    field['controls'] = sorted(field['controls'])
                    fields.append(field)
            elif node.tag==qn('w:instrText') and stack:
                stack[-1]['code'] += node.text or ''
                stack[-1]['controls'].add(run_display(document, p._p, node.getparent()))
            elif node.tag==qn('w:t'):
                text = node.text or ''
                if text and stack and stack[-1]['result']:
                    stack[-1]['text'] += text
                    display = run_display(document, p._p, node.getparent())
                    if display not in stack[-1]['displays']:
                        stack[-1]['displays'].append(display)
                offset += len(text)
        numbering = sequence_number(effective_numbering(document, p._p), counters)
        spans=[];active=None;offset=0
        for node in p._p:
            if node.tag==qn('w:bookmarkStart') and node.get(qn('w:name'),'').startswith('WFMCite_'):
                origin=re.fullmatch(r'WFMCite_(\d+)_\d+',node.get(qn('w:name'),''))
                active={'id':node.get(qn('w:id')),'start':offset,'text':'','segments':[],
                        'origin':int(origin[1]) if origin else None}
            elif node.tag==qn('w:bookmarkEnd') and active and node.get(qn('w:id'))==active['id']:
                spans.append(active);active=None
            text=''.join(t.text or '' for t in node.iter(qn('w:t')))
            if active and text:
                active['segments'].append({'start':offset, 'end':offset+len(text),
                                           'display':run_display(document, p._p, node)})
                active['text'] += text
            offset+=len(text)
        paragraphs.append({'text':paragraph_text(p._p),'fields':fields,'numbering':numbering,'spans':spans})
    return {'paragraphs':paragraphs,'bookmarks':bookmarks,'reference_rows':reference_rows}


def verify_state(state,params):
    """Require matching REF codes, cached values and real bibliography targets."""
    rows=state['paragraphs'];bookmarks=state['bookmarks']
    from citation_workflow import paragraph_sha256, _citation_offset
    ref_rows=state['reference_rows']
    for p in params['placements']:
        spacing=p.get('spacing','none')
        expected=(' ' if spacing in {'space-before','spaces-around'} else '')+p['citation_text'].strip()+(' ' if spacing in {'space-after','spaces-around'} else '')
        selected=[]
        for row in rows:
            for span in row['spans']:
                if span['origin']!=p['paragraph_index'] or span['text']!=expected:continue
                original=row['text'][:span['start']]+row['text'][span['start']+len(span['text']):]
                if paragraph_sha256(original)==p['paragraph_sha256'] and _citation_offset(original,p)==span['start']:
                    selected.append((row, span))
        if len(selected)!=1:return False
        row, span = selected[0]
        citation_text = p['citation_text'].strip()
        try:numbers,tokens=numeric_tokens(citation_text)
        except ValueError:return False
        if numbers!=[int(r.split('-')[-1]) for r in p['reference_ids']]:return False
        start = span['start'] + int(spacing in {'space-before', 'spaces-around'})
        end = start + len(citation_text)
        expected_display = p.get('display', 'inline')
        if any(segment['display'] != expected_display for segment in span['segments']
               if segment['start'] < end and segment['end'] > start):
            return False
        for token in tokens:
            number=int(token[1])
            if not 1<=number<=len(ref_rows):return False
            target=ref_rows[number-1]
            if not rows[target]['numbering'] or rows[target]['numbering']['format']!='decimal' or rows[target]['numbering']['value']!=number:return False
            found=False
            for field in row['fields']:
                if (field['start'] != start + token.start() or field['displays'] != [expected_display]
                        or field['controls'] != [expected_display]):
                    continue
                code=re.fullmatch(r'\s*REF\s+(\w+)\s+(.*?)\s*',field['code'])
                if not code:continue
                switches=code[2].split()
                if switches[-2:]==[r'\*','MERGEFORMAT']:switches=switches[:-2]
                elif 'display' in p:continue
                if sorted(switches)!=[r'\h',r'\n']:continue
                ends=bookmarks.get(code[1],[])
                if len(ends)==1 and ends[0]=={'paragraph':target,'paired':True} and field['text']==f'[{number}]':found=True
            if not found:return False
    return True


def snapshot(index,params,baseline=None):
    """Mask only validated newly generated citation spans/reference numbering."""
    from citation_workflow import extract_citation_data
    from .inspect import paragraph_text

    document=index.document
    if baseline is None:
        _,refs=extract_citation_data(document)
        reference_nodes={document.paragraphs[r['paragraph_index']]._p for r in refs}
        old_names={n.get(qn('w:name')) for n in document.element.iter(qn('w:bookmarkStart'))}
        target_nodes={document.paragraphs[p['paragraph_index']]._p for p in params['placements']}
    else:reference_nodes,old_names,target_nodes=baseline
    paragraphs=[]
    for p in document.element.body.iter(qn('w:p')):
        clone=copy.deepcopy(p)
        removing=None
        for child in list(clone):
            if child.tag==qn('w:bookmarkStart'):
                name=child.get(qn('w:name'),'')
                if p in target_nodes and name.startswith('WFMCite_') and name not in old_names:
                    removing=child.get(qn('w:id'));clone.remove(child);continue
                if p in reference_nodes and name.startswith('WFMRef_') and name not in old_names:
                    ident=child.get(qn('w:id'));clone.remove(child)
                    for end in list(clone.findall(qn('w:bookmarkEnd'))):
                        if end.get(qn('w:id'))==ident:clone.remove(end)
                    continue
            if removing is not None:
                if child.tag==qn('w:bookmarkEnd') and child.get(qn('w:id'))==removing:removing=None
                clone.remove(child)
        if p in reference_nodes:
            from .properties import PARAGRAPH_TAGS, read_paragraph
            effective={k:read_paragraph(document,p,k) for k in PARAGRAPH_TAGS if k!='numbering_suppressed'}
            pr=clone.find(qn('w:pPr'))
            if pr is not None:
                tags={qn('w:'+tag) for tagset in PARAGRAPH_TAGS.values() for tag in tagset}
                for child in list(pr):
                    if child.tag in tags:pr.remove(child)
                if len(pr)==0 and not pr.attrib:clone.remove(pr)
            prefix=PREFIX.match(paragraph_text(clone))
            if prefix:strip_text_prefix(clone,prefix.end())
        tokens=[]
        for child in clone:
            if child.tag==qn('w:r'):
                pr=child.find(qn('w:rPr'))
                props=xml_value(pr) if pr is not None and len(pr) else ''
                for node in child:
                    if node.tag==qn('w:rPr'):continue
                    if node.tag==qn('w:t'):tokens.extend((c,props) for c in node.text or '')
                    else:tokens.append(xml_value(node))
            else:tokens.append(xml_value(child))
        if p in reference_nodes:tokens.append(('effective_reference_paragraph',effective))
        paragraphs.append((p,tokens))
    # Existing numbering must be byte-for-byte unchanged; only append definitions.
    numbering=next((p.element for p in document.part.package.parts if str(p.partname)=='/word/numbering.xml'),[])
    if baseline is None:
        old_numbering=[xml_value(n) for n in numbering]
    else:old_numbering=None
    parts={str(p.partname):p.blob for p in document.part.package.parts
           if str(p.partname) not in {'/word/document.xml','/word/numbering.xml'}}
    tree=copy.deepcopy(document.element)
    for p in list(tree.iter(qn('w:p'))):p.getparent().replace(p,OxmlElement('w:p'))
    parts['document_nonparagraph_structure']=xml_value(tree)
    return paragraphs,parts,(reference_nodes,old_names,target_nodes),old_numbering


def verify_preservation(index,params,before):
    if not verify_state(read_state(index.document),params):
        raise FormatError('Citation fields, display or bookmarks do not match the request')
    after=snapshot(index,params,before[2])
    if before[:2]!=after[:2]:
        raise FormatError('Citation changed unrelated content, objects or formatting')
    old=before[3];numbering=index.document.part.numbering_part.element
    if [xml_value(n) for n in list(numbering)[:len(old)]]!=old:
        raise FormatError('Citation changed an existing numbering definition')
