# CAD Automation with Python + SolidWorks

This project automates **3D CAD modeling** in SolidWorks using Python and the COM API.

## Features
- Generate parametric cylinders from user input (`main.py`)
- Batch model creation from CSV (`batch_cylinder.py`)
- Export both **SLDPRT** and **STL** for 3D printing

## Requirements
- Python 3.x
- SolidWorks 2023 (with COM enabled)
- [pywin32](https://pypi.org/project/pywin32/)

Install dependency:
```bash
pip install pywin32
