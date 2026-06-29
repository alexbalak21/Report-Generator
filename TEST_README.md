# Test README

This file explains how to run the repository tests and how the test suite behaves when optional dependencies are missing.

## Setup

1. Create and activate the virtual environment:

```powershell
python -m venv env
.\env\Scripts\Activate.ps1
```

2. Install the project dependencies:

```powershell
pip install -r requirements.txt
```

## Running tests

From the repository root:

```powershell
python -m unittest discover -v
```

This command will discover and execute the tests under the `tests/` directory.

## Dependency notes

The repository tests require the following packages:

- `openpyxl`
- `python-docx`

If either package is unavailable, some tests are skipped gracefully rather than failing:

- `tests/test_report_generator.py` skips when `openpyxl`, `python-docx`, or report generator imports are unavailable.
- `tests/test_word_processor.py` skips when `python-docx` or `WordProcessor` cannot be imported.

The remaining processor tests in `tests/test_processors.py` do not require external Word or Excel packages.

## Expected test command output

When the environment is set up correctly, the command should run all available tests and report any failures.

If dependencies are missing, you may see skipped tests, but the `tests/test_processors.py` suite should still execute successfully.
