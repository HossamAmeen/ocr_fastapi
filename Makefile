PYTHON ?= python3
VENV := venv
PIP := $(VENV)/bin/pip
PY := $(VENV)/bin/python

.PHONY: venv install run clean

venv:
	$(PYTHON) -m venv $(VENV)
	$(PIP) install --upgrade pip

install: venv
	$(PIP) install -r requirements.txt

run-web:
	$(VENV)/bin/uvicorn app.main:app --reload --host 0.0.0.0 --port 8000

run:
	python desktop.py

build-exe:
	pyinstaller --onefile --windowed --name="OCRExcelGenerator" desktop.py

clean:
	rm -rf $(VENV) build dist *.spec uploads/* outputs/* app/__pycache__ app/*/__pycache__

