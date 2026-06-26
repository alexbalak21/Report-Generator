import json

class MappingLoader:
    def __init__(self, mapping_path):
        self.mapping_path = mapping_path

    def load(self):
        with open(self.mapping_path, "r", encoding="utf-8") as f:
            return json.load(f)
