# Excel–Python–VBA Automation: Retail Financial Processing Tool

## Overview
This project demonstrates the integration of Excel, Python, and VBA to automate financial data processing. Instead of performing manual calculations in Excel, a VBA macro triggers a Python script using xlwings, which processes financial data and writes computed results back into Excel automatically.

## Problem Statement
Retail businesses often require repetitive financial analysis, including:

- Calculating Gross Profit
- Estimating Tax Liability
- Computing Net Profit

Such calculations become inefficient and error-prone when done manually, especially across multiple records.  
This project automates the process by allowing Excel to call Python, perform calculations programmatically, and return results seamlessly into Excel.

## Dataset
The dataset used in this project is synthetic but realistic, designed to resemble retail financial performance data.

### Key Features in the Dataset
- Store Name – Store identifier  
- Line of Business – Business category  
- Total Revenue – Annual revenue  
- Total Expenses – Annual expenses  

### Additional Fields Computed by Python
- Gross Profit  
- Tax Percentage
- Tax Amount
- Net Profit

## Automation Workflow
The workflow includes:

- Reading data from the Input sheet in Excel  
- Processing financial calculations in Python
- Returning the processed data into the Output sheet  
- Allowing execution through a simple VBA macro trigger

## Technologies Used
- Python
- Pandas
- Numpy
- xlwings
- Excel
- VBA

## Conclusion

This project demonstrates a practical automation workflow combining Excel, Python, and VBA. It showcases the ability to build real-world automation tools, enhance efficiency, and design user-friendly interfaces.

## Author
**Shitikshu Vyas**
