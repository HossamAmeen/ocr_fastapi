PYTHON ?= python3
VENV := .venv
PIP := $(VENV)/bin/pip
PY := $(VENV)/bin/python

.PHONY: venv install run clean

venv:
	$(PYTHON) -m venv $(VENV)
	$(PIP) install --upgrade pip

install: venv
	$(PIP) install -r requirements.txt

run:
	$(VENV)/bin/uvicorn app.main:app --reload --host 0.0.0.0 --port 8000

clean:
	rm -rf $(VENV) uploads/* outputs/* app/__pycache__ app/*/__pycache__
