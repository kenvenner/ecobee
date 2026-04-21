# Project Instructions for Claude

## Overview

This is a python 3.13+ command line project that leverages ecobee API (python-ecobee-api) to control thermostats and track temperatures in a rental property.

---

## Coding Standards
- Use **type hints** for all function parameters and return values.
- Follow **PEP 8** naming conventions.
- Use **f-strings** for string formatting.
- Avoid wildcard imports (`from module import *`).
- Keep functions under **40 lines** where possible.
- Use `logging` instead of `print` for debug/info output.

---

## Dependencies
- Python 3.13+
- `ruff` for linting
- `ruff` for formatting
- `pytest` for testing
- `ty` for type checking

---

## Commands
- \'python villaecobee.py >> run.log 2>>error.log\' - Run the temp capture and thermostat setting application
- \'python vcconvert2.py'\ - Run the xlsx to txt booking conversion tooling
