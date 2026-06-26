import re
from docx import Document

PLACEHOLDER_PATTERN = r"\{\{(.*?)\}\}"

class WordProcessor:
    def __init__(self, filepath):
        self.filepath = filepath
        self.document = Document(filepath)

    def extract_placeholders(self):
        """Return all placeholders {{...}} found in the document."""
        text = "\n".join([p.text for p in self.document.paragraphs])
        return re.findall(PLACEHOLDER_PATTERN, text)

    def fill_placeholders(self, mapping, output_path):
        """
        Replace placeholders using mapping:
        { "date": "2024-06-12", ... }
        """
        for paragraph in self.document.paragraphs:
            for key, value in mapping.items():
                placeholder = f"{{{{{key}}}}}"
                if placeholder in paragraph.text:
                    paragraph.text = paragraph.text.replace(placeholder, str(value))

        self.document.save(output_path)
        