from pathlib import Path

path = Path(r'S:\DEV\Report-Generator\app\gui\config_persistence.py')
text = path.read_text(encoding='utf-8')
needle = "def load_mapping_config(mapping_path: str) -> dict:\n" \
         "    \"\"\"\n" \
         "    Read the 'config' block from a mapping JSON.\n" \
         "    Returns an empty dict on any error.\n" \
         "    \"\"\"\n" \
         "    try:\n" \
         "        return MappingLoader(mapping_path).load_config()\n" \
         "    except Exception:\n" \
         "        return {}\n"

if needle not in text:
    raise RuntimeError('Needle not found in config_persistence.py')

replacement = needle + "\n" \
              "def list_mapping_paths() -> list[str]:\n" \
              "    \"\"\"Return the list of known mapping file paths from the registry.\"\"\"\n" \
              "    try:\n" \
              "        return [path for _, path in mapping_list()]\n" \
              "    except Exception:\n" \
              "        return []\n"

text = text.replace(needle, replacement, 1)
path.write_text(text, encoding='utf-8')
print('patched config_persistence')
