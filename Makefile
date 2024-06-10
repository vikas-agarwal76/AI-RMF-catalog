# Makefile

.ONESHELL:
SHELL := /bin/bash

SOURCE_INIT = ./venv

all: run

run: transform

transform:
	source $(SOURCE_INIT)/bin/activate;
	python scripts/xlsx_to_oscal_catalog.py