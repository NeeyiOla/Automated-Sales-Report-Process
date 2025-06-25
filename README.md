# Automated-Sales-Report-Process

## Automated Weekly Sales Workflow & Forecasting using Power Automate, Power BI, and Excel 

![Waggle Lapdog and Lapcat Logos](asset/Images/waggle_logo_yellow.png)

# Table of Content
- [Project Title](#project-title)
- [My Role](#my-role)
- [Project Overvies](#project-overview)
- [Problem Statement](#problem-statement)
- [Stakeholder Engagement](#stakeholder-engagement)
  - [Target stakeholder](#target-stakeholder)
  - [Use Cases](#use-cases)
  - [Stakeholder Stories](#stakeholder-stories)
  - [Acceptance-Criteria](#acceptance-criteria)
  - [Success Criteria](#success-criteria)
- [Data Source](#data-source)
  - [Dataset](#dataset)
  - [Data model stucture](#data-model-structure)
- [Methodolody](#methodology)
  - [Tool Used](#tool-used)
  - [Development](#development)
    - [Project Planning & Requirement gathering](#project-planning--requirement-gathering)
    - [Data Exploration & Profiling](#data-exploration--profiling)
    - [Automation Workflow_Using Power Automate OneDrive and SharePoint](#automation_workflow_using_power_automate,_onedrive_and_sharepoint)
    - [ETL Process Using Power Query and Data modeling](#etl-process-using-power-query-and-data-modeling)
    - [Measures Developmenr Using DAX](#measures-development-using-dax)
    - [Dashboard Design & Visualisation](#dashboard-design--visualisation)
    - [Plushing and Collaboration](#publishing-and-collaboration)
    - [Documentation & Version control](#documentation--version-control)
    - [Review & Iteration](#review--iteration)
- [Detailed Insights and Recommendation](#detailed-insights-and-recommendations)
  - [CEO Executive Summary](#ceo-executive-summary)
- [P&L Statement and Financial Forecasting Model](#p&l_statement_and_financial_forecasting_model)
 

# Project Title
Automated Weekly Sales Workflow & Forecasting using POwer Automate, Power BI, and Excel.

# My Role
Business Intelligence Analyst & Workflow Automation Developer 

# Project Overview
This project automates the weekly sales data workflow, integrates historical data collection, and delivers real-time financial reporting and forecasting using Power Automate, Power BI, and Excel. 

# Problem Statement 
Manual handling of weekly sales reports created inefficiencies, inconsistency in reporting, and delays in decision-making. The lack of automation and centralised data limited the ability to conduct trend analysis, forecasting, and performance tracking.

# Stakeholder Engagement 
Weekly sync with the Sales Team, Finance Department, and C-Suite to identify reporting needs and automate manual tasks. Workshops were held to to define KPIs and success measures.

#Target Stakeholders 
- Sale Managers
- Finacial Analyst
- Operations Director
- C-Level Excutive (CEO, CFO)

# Use Cases 
- Automatic detection and transfer of weekly Excel reports
- Data consolidation and historical archiving
- Real-time Power BI updates
- Alert notifications for performance thresholds
- Financial forecasting and variance detection

# Stakeholder Stories
- "As a Sale Manager, I want to be notified when weekly sales reports are ready so i can prepare my sales review."
- "As a Financial Analyst, I need consistenthistorical sale data to generate Monthly P&L reports"
- "As the CEO, I want real-time KPIs and forecast alerts to track performance against goals."


# Acceptance Criteria
- Sales files are detected and archived weekly or every 7days 
- Power BI dashboard updates automatically after each file upload
- Stakeholder recieve notification and alerts
- P&L forecast is generated based on YTD sales


# Success Criteria
- 100% automation of weekly file ingestion process
- Power BI report refresh time under 5 minutes post-upload
- Forecast Model generates accurate projections within 5% error margin


# Data Source
- Weekly Excel files from Sales Team (OneDrive)
- Hosted in SharePoint: /Shared Document/Weekly Sales Upload and /Historical_Sales

### **DATASET**  
Each file includes:  
- Date, Product, Region
- Unit Sold, Unit Price, Revenue
- Unit Cost, Admin Cost, Marketing Cost


### **DATA MODEL STRUCTURE**
- Fact Table: Weekly Sales
- Dimension Tables: Date (Time Intelligence)
- Calculated Columns: Total Cost, Gross Profit, Profit Margin, Cost per Unit


# Methodology 
Agile developmenmt with bi-weekly sprints, stakeholder demos, and flow testing.  

### Tool Used 
- Power Automate
- Power BI Desktop & Service
- Execl (for P&L and forecasting
- SharePoint and OneDrive for Business
- DAX & Power Query

### Development 

##### *Project Planning & and Requirement Gathering* 
Collaborated with Stakeholders to define the reporting cadence, file structure, and KPI expectations.



##### *Automation Workflow Using Power Automate, OneDrive and SharePoint*

- Sales team Uploads weekly Excel (more like every 7days sales) to: /GitHub/ Power Automate/Weekly Sale Data Upload (OneDrive for Business)
- Power Automate detects the upload and:
  - Moves the files to /Shared Documents/Weekly Sales Upload/ (SharePoint)
  - Delete the file from OneDrive to prevent clustter
  - Copies the files to an archival folder: /Historical_Sales/ (SharePoint)
  - Sends an email notification to stakeholders tat a new file has been processed
- Power BI is connected to the Historical_Sales folder using the SharePoint folder connector
- This enables automated appending of weekly data and dashboard refresh

##### *Data Exploration & Profiling*
Performed integrity checks and missing value analysis using Power Query. Identified outliers and ensured consistency.

##### *ETL Process Using Power Query and Data Modeling* 
- Used the Folder connector to combine Excel files
- Cleaned and transformed fields: Dates, Costs, Revenue
- Related fact and dimensiontables for robust analysis

##### *Measures Development Using DAX*
Created measures: 
- Total Revenue, Total COGS, Total Cost
- Gross Profit, Profit Margin %, Average Gross profit
- COGS MOM %, COGS MoM Indicator (Color-coded)
- Total Unit Sold, Cost per Unit

##### *Dashboard Design & Visualisation*
- KPI cards, Bar charts, Line trends
- Drill-through from Product > Region > Unit Sold Monthly Trend
- Tooltip page
- Profitability scorecard and alert centre

##### *Publishing and Collaboration*
- Publish reports to Power BI Service
- Shared with stakeholders via Power BI apps and automated alerts.
- Collaboration on Outlook email, Microsoft Team

##### *Documentation & Version Control
Used GitHub for project files, Power BI versions, and flow documentation.  
Markdown for README.

##### *Review & Iteration*
Incorperated stakeholder feedback, Adjusted alerts, added new KPIs, improved visuals and usability.

# Detailed Insights and Recommendation.



# P&L Statement and Financial Forecasting Model




