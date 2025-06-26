# Automated-Sales-Report-Process

## Automated Weekly Sales Workflow & Forecasting using Power Automate, Power BI, and Excel 

![Automated Sales Reprt](Asset/Images/Cover_Page.png)

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

#### *Project Planning & and Requirement Gathering* 
Collaborated with Stakeholders to define the reporting cadence, file structure, and KPI expectations.



#### *Automation Workflow Using Power Automate, OneDrive and SharePoint*

- Sales team Uploads weekly Excel (more like every 7days sales) to: /GitHub/ Power Automate/Weekly Sale Data Upload (OneDrive for Business)
- Power Automate detects the upload and:
  - Moves the files to /Shared Documents/Weekly Sales Upload/ (SharePoint)
  - Delete the file from OneDrive to prevent clustter
  - Copies the files to an archival folder: /Historical_Sales/ (SharePoint)
  - Sends an email notification to stakeholders tat a new file has been processed
- Power BI is connected to the Historical_Sales folder using the SharePoint folder connector
- This enables automated appending of weekly data and dashboard refresh

![Dataflow Process and Report](Asset/Images/Dataflow_Process_and_Report.png)

#### *Data Exploration & Profiling*
Performed integrity checks and missing value analysis using Power Query. Identified outliers and ensured consistency.

#### *ETL Process Using Power Query and Data Modeling* 
- Used the Folder connector to combine Excel files
- Cleaned and transformed fields: Dates, Costs, Revenue
- Related fact and dimensiontables for robust analysis

#### *Measures Development Using DAX*
Created measures: 
- Total Revenue, Total COGS, Total Cost
- Gross Profit, Profit Margin %, Average Gross profit
- COGS MOM %, COGS MoM Indicator (Color-coded)
- Total Unit Sold, Cost per Unit

**KPIs DAX MEASURES**  
```Dax
Average Gross Profit = AVERAGE(Sales[Goss Profit])
```

```Dax
Gross Profit = [Total Revenue] - [Total Cost]
```

```Dax
Profit Margin % = DIVIDE([Gross Profit], [Total Revenue], 0)
```

```Dax
Total COGS = SUM(Sales[Unit Cost])
```

```Dax
Total Cost = 
SUMX('Sales',
 'Sales'[Unit Cost] * 'Sales'[Units Sold] + 'Sales'[Admin Cost] + 'Sales'[Marketing Cost])
```

```Dax
Total Revenue = SUM('Sales'[Revenue])
```

```Dax
Total Unit Sold = SUM('Sales'[Units Sold])
```


**MONTH OVER MONTH DAX MEASURES**  

```DAX
COGS MoM % = 
VAR SelectedMonth = MAX('Date'[Month])
VAR SelectedYear = MAX('Date'[Year])

VAR PreviousMnth = IF(SelectedMonth = 1, 12, SelectedMonth - 1)
VAR PreviousYr = IF(SelectedMonth = 1, SelectedYear - 1, SelectedYear)

VAR CurrentCOGS =
    CALCULATE(
        [Total COGS],
        'Date'[Month] = SelectedMonth,
        'Date'[Year] = SelectedYear
    )

VAR PreviousCOGS =
    CALCULATE(
        [Total COGS],
        'Date'[Month] = PreviousMnth,
        'Date'[Year] = PreviousYr,
        REMOVEFILTERS('Date')
    )

RETURN
    IF(
        PreviousCOGS <> 0,
        DIVIDE(CurrentCOGS - PreviousCOGS, PreviousCOGS),
        BLANK()
    )
```

```DAX
COGS MoM Indicator = 
VAR Change = [COGS MoM %]
VAR Arrow = 
    SWITCH(
        TRUE(),
        Change > 0, "▲",
        Change < 0, "▼",
        BLANK()
    )
VAR FormattedValue = FORMAT(ABS(Change), "0.0%")
RETURN
    IF(NOT ISBLANK(Change), Arrow & " " & FormattedValue)

```

```DAX
COGS MoM Δ = 
VAR SelectedMonth = MAX('Date'[Month])
VAR SelectedYear = MAX('Date'[Year])

VAR PreviousMnth = IF(SelectedMonth = 1, 12, SelectedMonth - 1)
VAR PreviousYr = IF(SelectedMonth = 1, SelectedYear - 1, SelectedYear)

VAR CurrentCOGS =
    CALCULATE(
        [Total COGS],
        'Date'[Month] = SelectedMonth,
        'Date'[Year] = SelectedYear
    )

VAR PreviousCOGS =
    CALCULATE(
        [Total COGS],
        'Date'[Month] = PreviousMnth,
        'Date'[Year] = PreviousYr,
        REMOVEFILTERS('Date')
    )

RETURN
    IF(
        NOT ISBLANK(PreviousCOGS),
        CurrentCOGS - PreviousCOGS
    )
```

```DAX
Gross Profit MoM % = 
VAR SelectedMonth = MAX('Date'[Month])
VAR SelectedYear = MAX('Date'[Year])

VAR PreviousMnth = IF(SelectedMonth = 1, 12, SelectedMonth - 1)
VAR PreviousYr = IF(SelectedMonth = 1, SelectedYear - 1, SelectedYear)

VAR CurrentGrossProfit =
    CALCULATE(
        [Gross Profit],
        'Date'[Month] = SelectedMonth,
        'Date'[Year] = SelectedYear
    )

VAR PreviousGrossProfit =
    CALCULATE(
        [Gross Profit],
        'Date'[Month] = PreviousMnth,
        'Date'[Year] = PreviousYr,
        REMOVEFILTERS('Date')
    )

RETURN
    IF(
        PreviousGrossProfit <> 0,
        DIVIDE(CurrentGrossProfit - PreviousGrossProfit, PreviousGrossProfit),
        BLANK()
    )
```

```DAX
Gross Profit MoM Indicator = 
VAR Change = [Gross Profit MoM %]
VAR Arrow = 
    SWITCH(
        TRUE(),
        Change > 0, "▲",
        Change < 0, "▼",
        BLANK()
    )
VAR FormattedValue = FORMAT(ABS(Change), "0.0%")
RETURN
    IF(NOT ISBLANK(Change), Arrow & " " & FormattedValue)

```

```DAX
Gross Profit MoM Δ = 
VAR SelectedMonth = MAX('Date'[Month])
VAR SelectedYear = MAX('Date'[Year])

VAR PreviousMnth = IF(SelectedMonth = 1, 12, SelectedMonth - 1)
VAR PreviousYr = IF(SelectedMonth = 1, SelectedYear - 1, SelectedYear)

VAR CurrentGP =
    CALCULATE(
        [Gross Profit],
        'Date'[Month] = SelectedMonth,
        'Date'[Year] = SelectedYear
    )

VAR PreviousGP =
    CALCULATE(
        [Gross Profit],
        'Date'[Month] = PreviousMnth,
        'Date'[Year] = PreviousYr,
        REMOVEFILTERS('Date')
    )

RETURN
    IF(
        NOT ISBLANK(PreviousGP),
        CurrentGP - PreviousGP
    )
```

```DAX
Profit Margin MoM % = 
VAR SelectedMonth = MAX('Date'[Month])
VAR SelectedYear = MAX('Date'[Year])

VAR PreviousMnth = IF(SelectedMonth = 1, 12, SelectedMonth - 1)
VAR PreviousYr = IF(SelectedMonth = 1, SelectedYear - 1, SelectedYear)

VAR CurrentMargin =
    CALCULATE(
        [Profit Margin %],
        'Date'[Month] = SelectedMonth,
        'Date'[Year] = SelectedYear
    )

VAR PreviousMargin =
    CALCULATE(
        [Profit Margin %],
        'Date'[Month] = PreviousMnth,
        'Date'[Year] = PreviousYr,
        REMOVEFILTERS('Date')
    )

RETURN
    IF(
        NOT(ISBLANK(PreviousMargin)),
        CurrentMargin - PreviousMargin,
        BLANK()
    )
```

```DAX
Profit Margin MoM Indicator = 
VAR Change = [Profit Margin MoM %]
VAR Arrow = 
    SWITCH(
        TRUE(),
        Change > 0, "▲",
        Change < 0, "▼",
        BLANK()
    )
VAR FormattedValue = FORMAT(ABS(Change), "0.0%")
RETURN
    IF(NOT ISBLANK(Change), Arrow & " " & FormattedValue)
```

```DAX
Revenue MoM %%% = 
VAR SelectedMonth = MAX('Date'[Month])
VAR SelectedYear = MAX('Date'[Year])

VAR PreviousMnth =
    IF(SelectedMonth = 1, 12, SelectedMonth - 1)

VAR PreviousYr =
    IF(SelectedMonth = 1, SelectedYear - 1, SelectedYear)

VAR RevenueCurrent =
    CALCULATE(
        [Total Revenue],
        'Date'[Month] = SelectedMonth,
        'Date'[Year] = SelectedYear
    )

VAR RevenuePrevious =
    CALCULATE(
        [Total Revenue],
        'Date'[Month] = PreviousMnth,
        'Date'[Year] = PreviousYr,
        REMOVEFILTERS('Date')  // 👈 ensures previous month can be evaluated even when filtered
    )

RETURN
    IF(
        RevenuePrevious <> 0,
        DIVIDE(RevenueCurrent - RevenuePrevious, RevenuePrevious),
        BLANK()
    )
```

```DAX
Revenue MoM Indicator = 
VAR Change = [Revenue MoM %%%]
VAR Arrow = 
    SWITCH(
        TRUE(),
        Change > 0, "▲",
        Change < 0, "▼",
        BLANK()
    )

VAR FormattedValue = FORMAT(ABS(Change), "0.0%")

RETURN
    IF(
        NOT ISBLANK(Change),
        Arrow & " " & FormattedValue
    )
```

```DAX
Revenue MoM Δ = 
VAR SelectedMonth = MAX('Date'[Month])
VAR SelectedYear = MAX('Date'[Year])

VAR PreviousMnth = IF(SelectedMonth = 1, 12, SelectedMonth - 1)
VAR PreviousYr = IF(SelectedMonth = 1, SelectedYear - 1, SelectedYear)

VAR CurrentRevenue =
    CALCULATE(
        [Total Revenue],
        'Date'[Month] = SelectedMonth,
        'Date'[Year] = SelectedYear
    )

VAR PreviousRevenue =
    CALCULATE(
        [Total Revenue],
        'Date'[Month] = PreviousMnth,
        'Date'[Year] = PreviousYr,
        REMOVEFILTERS('Date')
    )

RETURN
    IF(
        NOT ISBLANK(PreviousRevenue),
        CurrentRevenue - PreviousRevenue
    )
```

```DAX
Total Cost MoM % = 
VAR SelectedMonth = MAX('Date'[Month])
VAR SelectedYear = MAX('Date'[Year])

VAR PreviousMnth = IF(SelectedMonth = 1, 12, SelectedMonth - 1)
VAR PreviousYr = IF(SelectedMonth = 1, SelectedYear - 1, SelectedYear)

VAR CurrentCost =
    CALCULATE(
        [Total Cost],
        'Date'[Month] = SelectedMonth,
        'Date'[Year] = SelectedYear
    )

VAR PreviousCost =
    CALCULATE(
        [Total Cost],
        'Date'[Month] = PreviousMnth,
        'Date'[Year] = PreviousYr,
        REMOVEFILTERS('Date')
    )

RETURN
    IF(
        PreviousCost <> 0,
        DIVIDE(CurrentCost - PreviousCost, PreviousCost),
        BLANK()
    )
```

```DAX
Total Cost MoM Indicator = 
VAR Change = [Total Cost MoM %]
VAR Arrow = 
    SWITCH(
        TRUE(),
        Change > 0, "▲",
        Change < 0, "▼",
        BLANK()
    )
VAR FormattedValue = FORMAT(ABS(Change), "0.0%")
RETURN
    IF(NOT ISBLANK(Change), Arrow & " " & FormattedValue)
```

```DAX
Total Cost MoM Δ = 
VAR SelectedMonth = MAX('Date'[Month])
VAR SelectedYear = MAX('Date'[Year])

VAR PreviousMnth = IF(SelectedMonth = 1, 12, SelectedMonth - 1)
VAR PreviousYr = IF(SelectedMonth = 1, SelectedYear - 1, SelectedYear)

VAR CurrentCost =
    CALCULATE(
        [Total Cost],
        'Date'[Month] = SelectedMonth,
        'Date'[Year] = SelectedYear
    )

VAR PreviousCost =
    CALCULATE(
        [Total Cost],
        'Date'[Month] = PreviousMnth,
        'Date'[Year] = PreviousYr,
        REMOVEFILTERS('Date')
    )

RETURN
    IF(
        NOT ISBLANK(PreviousCost),
        CurrentCost - PreviousCost
    )
```

```DAX
Unit Sold MoM % = 
VAR SelectedMonth = MAX('Date'[Month])
VAR SelectedYear = MAX('Date'[Year])

VAR PreviousMnth = IF(SelectedMonth = 1, 12, SelectedMonth - 1)
VAR PreviousYr = IF(SelectedMonth = 1, SelectedYear - 1, SelectedYear)

VAR CurrentCOGS =
    CALCULATE(
        [Total Unit Sold],
        'Date'[Month] = SelectedMonth,
        'Date'[Year] = SelectedYear
    )

VAR PreviousCOGS =
    CALCULATE(
        [Total Unit Sold],
        'Date'[Month] = PreviousMnth,
        'Date'[Year] = PreviousYr,
        REMOVEFILTERS('Date')
    )

RETURN
    IF(
        PreviousCOGS <> 0,
        DIVIDE(CurrentCOGS - PreviousCOGS, PreviousCOGS),
        BLANK()
    )
```


#### *Dashboard Design & Visualisation*
- KPI cards, Bar charts, Line trends
- Drill-through from Product > Region > Unit Sold Monthly Trend
- Tooltip page
- Profitability scorecard and alert centre

#### *Publishing and Collaboration*
- Publish reports to Power BI Service
- Shared with stakeholders via Power BI apps and automated alerts.
- Collaboration on Outlook email, Microsoft Team

#### *Documentation & Version Control*
Used GitHub for project files, Power BI versions, and flow documentation.  
Markdown for README.

#### *Review & Iteration*
Incorperated stakeholder feedback, Adjusted alerts, added new KPIs, improved visuals and usability.

# Detailed Insights and Recommendation.
## Sales Report Dashboard & Drill-Through Page

![Sales Report Dasboard](Asset/Images/Sales_Report_Dashboard.gif)

# P&L Statement and Financial Forecasting Model




