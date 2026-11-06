"""Audit regressions: real citation relationships, validation and module boundaries."""
import copy
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from dataclasses import replace

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / 'scripts'))
from citation_workflow import build_citation_task
from word_format import build_plan, apply_plan, verify_plan
from word_format.contracts import FormatError
from word_format.registry import REGISTRY
from word_format.properties import read_paragraph, PARAGRAPH_TAGS
from word_format.profiles import profile_plan, template_profile


class DynamicCitationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.source = self.folder/'source.docx'
        self.output = self.folder/'output.docx'
        doc = Document()
        doc.add_heading('绪论', 1)
        doc.add_paragraph('数据库索引改善查询性能。')
        doc.add_paragraph('另外一项研究。')
        doc.add_heading('参考文献', 1)
        doc.add_paragraph('[1] Author. Database[M]. 2020.')
        doc.add_paragraph('[2] Author. Index[J]. 2021.')
        doc.add_paragraph('致谢')
        doc.add_paragraph('感谢各位。')
        doc.save(self.source)

    def request(self, text='[1]', ids=None, position='before-terminal-punctuation', index=1):
        task = build_citation_task(Document(self.source))
        p = next(p for p in task['paragraphs'] if p['paragraph_index'] == index)
        return {'operations':[{'action':'citation.insert','target':{'type':'document'},'params':{'placements':[
            {'paragraph_index':index,'paragraph_sha256':p['paragraph_sha256'],
             'reference_ids':ids or ['ref-1'],'citation_text':text,'insert':position}]}}]}

    def execute(self, request):
        plan=build_plan(self.source,self.output,request)
        report=apply_plan(plan)
        self.assertEqual(verify_plan(plan,report)['status'],'passed')
        return Document(self.output), plan, report

    def test_extraction_stops_at_plain_acknowledgements_and_appendix(self):
        for marker in ['致谢','附录 A','Acknowledgments','Appendix B']:
            doc=Document(self.source);doc.paragraphs[6].text=marker;doc.save(self.source)
            task=build_citation_task(doc)
            self.assertEqual(len(task['references']),2)
            self.assertNotIn(7,[p['paragraph_index'] for p in task['paragraphs']])

    def test_dynamic_ref_bookmarks_native_numbering_and_original_text(self):
        doc,_,_=self.execute(self.request())
        self.assertEqual(doc.paragraphs[1].text,'数据库索引改善查询性能[1]。')
        codes=[n.text for n in doc.paragraphs[1]._p.iter(qn('w:instrText'))]
        self.assertTrue(any(c.strip().startswith('REF ') for c in codes))
        self.assertIn(r'\* MERGEFORMAT', ''.join(codes))
        bookmark=doc.paragraphs[4]._p.find(qn('w:bookmarkStart'))
        self.assertIsNotNone(bookmark)
        self.assertIn(bookmark.get(qn('w:name')), ''.join(codes))
        from word_format.citations import read_state
        self.assertEqual(read_state(doc)['paragraphs'][1]['fields'][0]['text'],'[1]')
        self.assertTrue(doc.paragraphs[4]._p.xpath('./w:pPr/w:numPr/w:numId'))
        self.assertEqual(doc.paragraphs[4].text,'Author. Database[M]. 2020.')
        self.assertEqual(doc.paragraphs[7].text,'感谢各位。')

    def test_citation_display_is_independent_of_position_and_first_run(self):
        doc = Document(self.source)
        doc.paragraphs[1].runs[0].font.superscript = True
        doc.save(self.source)
        request = self.request()
        placements = request['operations'][0]['params']['placements']
        raised = self.request(position='after-terminal-punctuation', index=2)
        placements.append({**raised['operations'][0]['params']['placements'][0],
                           'display': 'superscript'})
        output, _, _ = self.execute(request)
        for index, suffix, expected in [(1, '[1]。', False), (2, '。[1]', True)]:
            paragraph = output.paragraphs[index]
            result = next(run for run in paragraph.runs if run.text == '[1]')
            self.assertTrue(paragraph.text.endswith(suffix))
            self.assertIs(result.font.superscript, expected)
        self.assertIs(output.paragraphs[1].runs[0].font.superscript, True)

    def test_verifier_rejects_wrong_citation_display(self):
        request = self.request()
        request['operations'][0]['params']['placements'][0]['display'] = 'superscript'
        cap = REGISTRY['citation.insert']

        def wrong(index, ref, params):
            cap.handler(index, ref, params)
            result = next(r for r in index.document.paragraphs[1].runs if r.text == '[1]')
            result.font.superscript = False

        with patch.dict(REGISTRY, {'citation.insert': replace(cap, handler=wrong)}):
            with self.assertRaisesRegex(FormatError, 'Citation fields, display or bookmarks'):
                self.execute(request)
        self.assertFalse(self.output.exists())

    def test_sample_superscript_rule_applies_only_to_recognized_citations(self):
        sample = self.folder / 'sample.docx'
        doc = Document(self.source)
        style = doc.styles.add_style('Citation Mark', WD_STYLE_TYPE.CHARACTER)
        style.font.superscript = True
        doc.paragraphs[1].add_run('[1]', style)
        doc.save(sample)

        profile = template_profile(sample)
        rule = next((r for r in profile['rules'] if r['action'] == 'citation.display.set'), None)
        self.assertIsNotNone(rule)
        self.assertEqual(rule['params']['value'], 'superscript')
        self.assertFalse(any('vertAlign' in issue.get('properties', []) for issue in profile['issues']))

        doc = Document(self.source)
        paragraph = doc.paragraphs[1]
        paragraph.add_run('[1]与[2]支持结论，未列出的编号[99]保留。').bold = True
        paragraph.add_run('2').font.superscript = True
        doc.save(self.source)
        resolutions = {issue['id']: {'decision': 'ignore', 'reason': 'Unrelated sample settings'}
                       for issue in profile['issues']}
        plan = profile_plan(self.source, self.output, profile, resolutions=resolutions,
                            match_structure=True)
        report = apply_plan(plan)
        self.assertEqual(verify_plan(plan, report)['status'], 'passed')
        paragraph = Document(self.output).paragraphs[1]
        self.assertEqual([r.text for r in paragraph.runs if r.font.superscript], ['[1]', '[2]', '2'])
        self.assertEqual(paragraph.text, doc.paragraphs[1].text)

    def test_mixed_sample_keeps_both_displays_as_separate_evidence(self):
        doc = Document(self.source)
        doc.paragraphs[1].add_run('[1]').font.superscript = True
        doc.paragraphs[2].add_run('[2]')
        doc.save(self.source)
        task = build_citation_task(doc)
        self.assertEqual([[m['display'] for m in p['citations']] for p in task['paragraphs']],
                         [['superscript'], ['inline']])
        profile = template_profile(self.source)
        issue = next((i for i in profile['issues'] if i.get('property') == 'citation_display'), None)
        self.assertIsNotNone(issue)
        self.assertEqual({r['params']['value'] for r in issue['candidates']}, {'inline', 'superscript'})
        self.assertFalse(any(r['action'] == 'citation.display.set' for r in profile['rules']))

    def test_existing_ref_can_be_reformatted_without_rewriting_field(self):
        output, _, _ = self.execute(self.request())
        self.source = self.output
        self.output = self.folder / 'raised.docx'
        original = output.paragraphs[1]
        request = {'operations': [{'action': 'citation.display.set',
                                  'target': {'type': 'paragraph', 'role': 'body', 'all': True},
                                  'params': {'value': 'superscript'}}]}
        output, _, _ = self.execute(request)
        paragraph = output.paragraphs[1]
        self.assertEqual(paragraph.text, original.text)
        self.assertEqual([n.text for n in paragraph._p.iter(qn('w:instrText'))],
                         [n.text for n in original._p.iter(qn('w:instrText'))])
        result = next(r for r in paragraph.runs if r.text == '[1]')
        self.assertIs(result.font.superscript, True)
        self.assertTrue(all(r.font.superscript for r in paragraph.runs
                            if r._r.find(qn('w:instrText')) is not None))
        profile = template_profile(self.output)
        self.assertFalse(any('vertAlign' in i.get('properties', []) for i in profile['issues']))

    def test_reference_field_format_does_not_touch_an_unrelated_field(self):
        doc = Document(self.source)
        paragraph = doc.paragraphs[1]
        for target, name in [(doc.paragraphs[4], 'reference'), (doc.paragraphs[2], 'unrelated')]:
            start = OxmlElement('w:bookmarkStart')
            start.set(qn('w:id'), str(len(name)))
            start.set(qn('w:name'), name)
            target._p.insert(0, start)
            end = OxmlElement('w:bookmarkEnd')
            end.set(qn('w:id'), str(len(name)))
            target._p.append(end)
            field = OxmlElement('w:fldSimple')
            field.set(qn('w:instr'), f'REF {name} \\h')
            paragraph.add_run('[')
            field.append(paragraph.add_run('1')._r)
            paragraph._p.append(field)
            paragraph.add_run(']')
        doc.save(self.source)

        request = {'operations': [{'action': 'citation.display.set',
                                  'target': {'type': 'paragraph', 'role': 'body', 'all': True},
                                  'params': {'value': 'superscript'}}]}
        output, _, _ = self.execute(request)
        fields = list(output.paragraphs[1]._p.iter(qn('w:fldSimple')))
        alignments = [field.find('.//' + qn('w:vertAlign')) for field in fields]
        self.assertEqual(alignments[0].get(qn('w:val')), 'superscript')
        self.assertIsNone(alignments[1])
        self.assertEqual([r.text for r in output.paragraphs[1].runs if r.font.superscript], ['[', ']'])

    def test_literal_citations_require_actual_numeric_reference_labels(self):
        doc = Document(self.source)
        doc.paragraphs[1].add_run('文献[8]与数组[1]。')
        for paragraph in doc.paragraphs[4:6]:
            paragraph.text = paragraph.text[4:]
        task = build_citation_task(doc)
        self.assertEqual(task['paragraphs'][0]['citations'], [])

        from word_format.numbering import effective_numbering
        for paragraph in doc.paragraphs[4:6]:
            paragraph.style = 'List Number'
        num = effective_numbering(doc, doc.paragraphs[4]._p)[3]
        override = OxmlElement('w:lvlOverride')
        override.set(qn('w:ilvl'), '0')
        start = OxmlElement('w:startOverride')
        start.set(qn('w:val'), '8')
        override.append(start)
        num.append(override)
        task = build_citation_task(doc)
        self.assertEqual([m['text'] for m in task['paragraphs'][0]['citations']], ['[8]'])

    def test_citation_task_offsets_use_the_exported_source_text(self):
        doc = Document(self.source)
        doc.paragraphs[1].text = '  测试引文[1]。'
        body = build_citation_task(doc)['paragraphs'][0]
        mark = body['citations'][0]
        self.assertEqual(body['text'][mark['start']:mark['end']], '[1]')

    def test_conflicting_citation_display_requests_are_rejected_during_planning(self):
        request = self.request()
        request['operations'].append({'action': 'citation.display.set',
                                      'target': {'type': 'paragraph', 'role': 'body', 'all': True},
                                      'params': {'value': 'superscript'}})
        with self.assertRaisesRegex(FormatError, 'Conflicting citation display'):
            build_plan(self.source, self.output, request)

    def test_verifier_checks_separators_inside_superscript_citation(self):
        request = self.request('[1], [2]', ['ref-1', 'ref-2'])
        request['operations'][0]['params']['placements'][0]['display'] = 'superscript'
        cap = REGISTRY['citation.insert']

        def wrong(index, ref, params):
            cap.handler(index, ref, params)
            separator = next(r for r in index.document.paragraphs[1].runs if r.text == ', ')
            separator.font.superscript = False

        with patch.dict(REGISTRY, {'citation.insert': replace(cap, handler=wrong)}):
            with self.assertRaisesRegex(FormatError, 'Citation fields, display or bookmarks'):
                self.execute(request)
        self.assertFalse(self.output.exists())

    def test_reject_display_id_mismatch_before_publish(self):
        with self.assertRaises((ValueError,FormatError)):
            self.execute(self.request('[999]'))
        self.assertFalse(self.output.exists())

    def test_multiple_reference_display(self):
        doc,_,_=self.execute(self.request('[1][2]',['ref-1','ref-2']))
        self.assertEqual(doc.paragraphs[1].text,'数据库索引改善查询性能[1][2]。')
        self.assertEqual(len(list(doc.paragraphs[1]._p.iter(qn('w:instrText')))),2)

    def test_compressed_range_requires_explicit_reference_list(self):
        with self.assertRaises(FormatError):self.execute(self.request('[1-2]',['ref-1','ref-2']))
        self.assertFalse(self.output.exists())

    def test_later_insertion_reuses_reference_bookmark(self):
        doc,_,_=self.execute(self.request())
        names=[n.get(qn('w:name')) for n in doc.paragraphs[4]._p.iter(qn('w:bookmarkStart'))]
        self.source=self.output;self.output=self.folder/'second.docx'
        doc,_,_=self.execute(self.request(index=2))
        self.assertEqual(names,[n.get(qn('w:name')) for n in doc.paragraphs[4]._p.iter(qn('w:bookmarkStart'))])

    def test_complex_anchor_is_rejected_without_losing_objects(self):
        request=self.request()
        doc=Document(self.source)
        node=OxmlElement('w:drawing');doc.paragraphs[1].runs[0]._r.append(node);doc.save(self.source)
        with self.assertRaises((FormatError,ValueError)):
            self.execute(request)
        self.assertFalse(self.output.exists())

    def test_blank_reference_paragraph_does_not_shift_reference_ids(self):
        doc=Document(self.source)
        blank=doc.add_paragraph()._p
        doc.paragraphs[5]._p.addprevious(blank)
        doc.save(self.source)
        self.execute(self.request('[2]',['ref-2']))

    def test_existing_toc_does_not_shift_raw_placement_indices(self):
        doc=Document(self.source)
        style=doc.styles.add_style('WFM TOC Field',1)
        paragraph=doc.add_paragraph('Cached TOC',style=style)
        doc.paragraphs[1]._p.addprevious(paragraph._p)
        doc.save(self.source)
        self.execute(self.request(index=2))

    def test_no_numbering_part_is_created_safely(self):
        doc=Document(self.source)
        rel=next(r for r in doc.part.rels.values() if r.reltype.endswith('/numbering'))
        doc.part.drop_rel(rel.rId);doc.save(self.source)
        self.execute(self.request())

    def test_native_reference_indentation_and_tabs_are_preserved(self):
        doc=Document(self.source)
        for i in [4,5]:
            doc.paragraphs[i].style='List Number'
            doc.paragraphs[i].text=doc.paragraphs[i].text[4:]
        doc.save(self.source)
        keys=['left_indent','hanging_indent','tabs']
        before={k:read_paragraph(doc,doc.paragraphs[4]._p,k) for k in keys}
        out,_,_=self.execute(self.request())
        self.assertEqual(before,{k:read_paragraph(out,out.paragraphs[4]._p,k) for k in keys})

    def test_split_retains_leading_spaces_xml_preserve(self):
        doc=Document(self.source);doc.paragraphs[1].text='  Original spaces.  ';doc.save(self.source)
        out,_,_=self.execute(self.request(position='start'))
        for node in out.paragraphs[1]._p.iter(qn('w:t')):
            if (node.text or '')[:1].isspace() or (node.text or '')[-1:].isspace():
                self.assertEqual(node.get(qn('xml:space')),'preserve')

    def test_second_workflow_with_shifted_source_indices(self):
        doc,_,_=self.execute(self.request())
        extra=doc.add_paragraph('新增正文。')
        doc.paragraphs[1]._p.addprevious(extra._p)
        self.source=self.folder/'shifted.docx';doc.save(self.source)
        self.output=self.folder/'second.docx'
        self.execute(self.request(index=1))

    def test_reference_sequence_shared_with_outside_list_rejected(self):
        doc=Document(self.source)
        for i in [4,5]:doc.paragraphs[i].style='List Number'
        doc.add_paragraph('Unrelated numbered item',style='List Number');doc.save(self.source)
        with self.assertRaises(FormatError):self.execute(self.request())
        self.assertFalse(self.output.exists())

    def test_verifier_rejects_plain_text_impostor(self):
        cap=REGISTRY['citation.insert']
        def fake(index,ref,params):
            p=index.document.paragraphs[1]
            p.add_run('[1]')
        with patch.dict(REGISTRY,{'citation.insert':replace(cap,handler=fake)}):
            with self.assertRaises(FormatError):self.execute(self.request())
        self.assertFalse(self.output.exists())

    def test_verifier_rejects_field_pointing_at_other_reference(self):
        cap=REGISTRY['citation.insert']
        def wrong(index,ref,params):
            cap.handler(index,ref,params)
            p=index.document.paragraphs[1]
            code=next(p._p.iter(qn('w:instrText')))
            code.text=' REF nonexistent_bookmark \\n \\h '
        with patch.dict(REGISTRY,{'citation.insert':replace(cap,handler=wrong)}):
            with self.assertRaises(FormatError):self.execute(self.request())
        self.assertFalse(self.output.exists())

    def test_verifier_rejects_unrequested_numeric_picture_switch(self):
        cap=REGISTRY['citation.insert']
        def wrong(index,ref,params):
            cap.handler(index,ref,params)
            code=next(index.document.paragraphs[1]._p.iter(qn('w:instrText')))
            code.text += ' \\# "BAD" '
        with patch.dict(REGISTRY,{'citation.insert':replace(cap,handler=wrong)}):
            with self.assertRaises(FormatError):self.execute(self.request())
        self.assertFalse(self.output.exists())


if __name__=='__main__':unittest.main()
