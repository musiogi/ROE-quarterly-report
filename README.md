# ROE-quarterly-report
This project implements an **automated workflow to generate Outlook draft emails for stakeholders across 21 European jurisdictions**, providing country-level ROE analysis with embedded charts and supervisory insights.

The repository is based on a Python script to analyse ROE data for significant institutions using publicly available data from the ECB Data Portal. 

## Data sources
The analysis is based exclusively on publicly available supervisory data for Significant Institutions at Country and aggregated level: https://data.ecb.europa.eu/data/datasets/SUP/SUP.Q.B01.W0._Z.I2003._T.SII._Z._Z._Z.PCT.C  

## Key Features
### Outlook connection
Mails will be directly saved in your outlook drafts.
### Quarterly alignment
All countries are compared using a common reference quarter (latest available SSM data).
### Confidentiality-aware logic
If data for the reference quarter is unavailable for a country, the script avoids performing the analysis and explicitly flags: “Data is unavailable due to confidentiality issues”.

## How to run
1. Download the dataset from this repository https://github.com/musiogi/ROE-quarterly-report/blob/main/ECB%20Data%20Portal%20long_ROE.xlsx
2. Update the file path in the script  
3. Run the Python script  

## Outputs
The script produces:
- country-level ROE analysis
- comparison with the SSM aggregate
- automated draft emails with embedded charts

## Example of output
1. With available data
   
<img width="1362" height="915" alt="image" src="https://github.com/user-attachments/assets/a9f3b06b-5a4d-4773-8400-042a04952a11" />

2. Without available data

<img width="1360" height="919" alt="image" src="https://github.com/user-attachments/assets/c70faebb-62c5-4e1c-b3ac-6e75024d17e2" />
