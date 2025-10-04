# ExcelExporter

`ExcelExporter` is a Python utility to export Excel ranges or charts to PDF or PNG using **LibreOffice headless**. It works on Linux and retains cell styles from Excel files. For chart exports, it detects the chart’s exact cell range and exports only that portion.

---

## Features

- Export **cell ranges** to PDF or PNG.
- Export **charts** to PDF or PNG.
- Export **entire sheets** to PDF or PNG.
- Automatically preserves styling within the selected range.
- Uses LibreOffice headless for fully compatible Excel rendering.
- Works with both Excel files and openpyxl `Workbook` objects.

---

## Installation

1. Install LibreOffice CLI tools:

```bash
sudo apt install libreoffice
sudo apt install poppler-utils  # for pdftoppm (PNG exports)
```

2. Install Python dependencies:

```bash
pip install openpyxl
```


---

## Usage

```python
from openpyxl import load_workbook, Workbook
from filebridge import ExcelExporter

# Load an existing Excel workbook
wb = load_workbook("example.xlsx")

# Initialize ExcelExporter
exporter = ExcelExporter(wb)

# ------------------- Range Exports -------------------
# Export a cell range to PDF
exporter.export_range_to_pdf(sheet_name="Sheet1", cell_range="B2:D10", output_path="range.pdf")

# Export a cell range to PNG
exporter.export_range_to_png(sheet_name="Sheet1", cell_range="B2:D10", output_path="range.png")

# ------------------- Sheet Exports -------------------
# Export the entire sheet to PDF
exporter.export_sheet_to_pdf(output_path="sheet.pdf")

# Export the entire sheet to PNG
exporter.export_sheet_to_png(output_path="sheet.png")

# ------------------- Chart Exports -------------------
# Export a specific chart to PDF
exporter.export_chart_to_pdf(sheet_name="Sheet1", chart_name="Sales Chart", output_path="chart.pdf")

# Export a specific chart to PNG
exporter.export_chart_to_png(sheet_name="Sheet1", chart_name="Sales Chart", output_path="chart.png")
```

---

## Notes

- For **range and chart exports**, a temporary workbook is created with only the relevant cells to minimize output size.
- PNG exports are performed via `LibreOffice -> PDF -> pdftoppm -> PNG`.
- Ensure chart names exactly match the titles defined in Excel.
- LibreOffice CLI does not directly crop charts or ranges; this approach ensures only the desired content is exported.

---

## License

MIT License
