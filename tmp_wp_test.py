from app.core.word_processor import WordProcessor
mapping = {
    '{{temperature_reception}}': '4°C',
    '{{date_rapport}}': '21/06/2026',
    '{{numéro_rapport}}': '260621-1'
}
wp = WordProcessor('data/rapport_template.docx')
wp.fill_placeholders(mapping, 'out/tmp_wp_test.docx')
print('saved out/tmp_wp_test.docx')
from app.core.word_processor import WordProcessor as WP2
wp2 = WP2('out/tmp_wp_test.docx')
print(sorted(set(wp2.extract_placeholders())))
