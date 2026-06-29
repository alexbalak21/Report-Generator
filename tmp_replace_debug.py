from app.core.word_processor import WordProcessor

wp = WordProcessor('out/test_output.docx')
for p in wp._all_paragraphs():
    if '{{temperature_reception}}' in p.text or '{{numéro_rapport}}' in p.text:
        print('PARAGRAPH', p.text)
        for i, r in enumerate(p.runs):
            print('RUN', i, repr(r.text), 'contains placeholder?', '{{temperature_reception}}' in r.text, '{{numéro_rapport}}' in r.text)
        wp._replace_in_paragraph(p, {'{{temperature_reception}}': '4°C', '{{numéro_rapport}}': '260621-1'})
        print('after paragraph', p.text)
        for i, r in enumerate(p.runs):
            print('RUN_AFTER', i, repr(r.text))
        print('-----')
