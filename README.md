# Rapport-Generator
A small Python tool to generate Word reports from Excel data.

Features
- Read Excel spreadsheets and map data into a report template
- Generate `.docx` reports using `python-docx` and project templates
- Lightweight GUI for choosing input/output files (optional)

Requirements
- Python 3.8+
- See `requirements.txt` for Python package dependencies

Installation
1. Create and activate a virtual environment:

	 ```powershell
	 python -m venv env
	 .\env\Scripts\Activate.ps1
	 ```

2. Install dependencies:

	 ```powershell
	 pip install -r requirements.txt
	 ```

Usage
- Run the GUI application:

	```powershell
	python app/app.py
	```

- Or run the provided script to inspect/generate reports:

	```powershell
	python inspect_report.py
	```

Testing

Run the test suite with `pytest`:

```powershell
pip install pytest
python -m pytest
```

Project structure
- `app/` — application package (GUI, core logic, mappings)
- `exemples/` — example files
- `tests/` — unit tests (e.g., `test_word_processor.py`)

Notes
- This project currently does not include an explicit license file.
	Contact the maintainer for licensing information.

