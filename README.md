# OSP-Sales-Split-Analyzer-and-Final-Location-Generator

This Streamlit app helps you analyze and visualize the OSP final location allocations and how customer sales were split across them.

## 🔧 How It Works

1. Upload:
   - 📁 **Sales History Excel** (with weekly volumes)
   - 📁 **OSP Final Location Excel** (with weekly final locations)

2. It calculates:
   - Total sales per SKU
   - Split % per location based on IPAG & SCH
   - Weekly views of how demand is distributed

3. Output:
   - Well-structured Excel report
   - One sheet for SKU-IPAG splits per week
   - One sheet for SKU-SCH splits per week

## 📤 Export Format

```text
| SKU | IPAG | OSP WK1 Final Location | CSF Split% SKU-IPAG WK1 |
|-----|------|------------------------|---------------------------|
| 123 | AB12 | FE                     | 35.2%                     |
