from app.core.word_processor import WordProcessor

wp = WordProcessor('out/test_output.docx')
for p in wp._all_paragraphs():
    if '{{temperature_reception}}' in p.text or '{{numéro_rapport}}' in p.text:
        print('PARAGRAPH', p.text)
        for i, r in enumerate(p.runs):
            print('RUN', i, repr(r.text))
        print('--- merging ---')
        wp._merge_split_placeholders(p)
        for i, r in enumerate(p.runs):
            print('RUN_AFTER', i, repr(r.text))
        print('========')
