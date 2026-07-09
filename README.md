# Automated Business Case & Financial Model Builder

An automated Command-Line Interface (CLI) tool designed to take basic financial inputs for technology implementation projects and generate client-ready Excel workbooks. The generated workbook features dynamic ROI, NPV, and payback period calculations alongside an embedded visual trend chart.

---

## 🚀 Features

* **Professional Financial Design:** Uses industry-standard formatting conventions (e.g., blue text for hardcoded inputs, black text for formulas).
* **Dynamic Calculations:** Generates native Excel formulas for Net Present Value (NPV), Return on Investment (ROI), and payback periods rather than just static numbers.
* **Automated Visuals:** Embeds a line chart tracking Net Cash Flow, Cumulative Cash Flow, and Cumulative Present Value (PV) over the project timeline.
* **Multi-Tab Architecture:** Organizes data seamlessly across three dedicated sheets: `Assumptions`, `Financial Model`, and `Notes`.

---

## 🛠️ Setup & Installation

### 1. Prerequisites
Ensure you have Python 3 installed on your machine.

### 2. Install Dependencies
This project requires the `openpyxl` library to create and style Excel files. Install it via your terminal:

```bash
pip install openpyxl