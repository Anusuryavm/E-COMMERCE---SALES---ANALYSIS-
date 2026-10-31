# 📊 E-Commerce Retail Analytics - Excel Capstone Project

## 📌 Project Overview
This project presents an end-to-end data analytics solution for an E-Commerce Retail dataset containing **2,000 sales transactions** generating a total revenue of **$2,170,101**. 

The primary objective was to perform comprehensive data cleaning, standardization, modeling via a **Star Schema**, and statistical analysis using Microsoft Excel. Interactive Pivot Tables and Dashboards were built to uncover actionable business insights across descriptive, diagnostic, predictive, and prescriptive analytical layers.

---

## 🎯 Key Objectives
- **Data Cleaning & Handling Missing Values:** Deduplicate data and treat missing values using statistical techniques (Mean/Average imputation).
- **Data Formatting & Standardization:** Clean and format raw text inputs across all dimension tables using `PROPER`, `TRIM`, and `CLEAN` functions.
- **Business Logic Calculations:** Accurately derive total transaction amounts considering ordered quantities, unit prices, and applied discounts.
- **Data Modeling:** Establish a relational **Star Schema** linking `Sales Fact` with `Customer Dim`, `Product Dim`, and `Store Dim`.
- **Lookup Automation:** Automate data retrieval across dimensions using modern dynamic formulas (`XLOOKUP` & `VLOOKUP`).
- **Dashboard & Storytelling:** Create visual dashboards and frame analytical insights to drive strategic decision-making.

---

## 🗂️ Data Architecture & Schema

The dataset is structured into a **Star Schema** model comprising **3 Dimension Tables** and **1 Fact Table**:

```
                  +-------------------+
                  |   Customer Dim    |
                  |-------------------|
                  | Customer ID (PK)  |
                  | Name, Age, Gender |
                  | City, State       |
                  | Loyalty Level     |
                  +---------+---------+
                            |
                            | 1:N
                            v
+-------------------+     +-------------------+     +-------------------+
|    Product Dim    |     |    Sales Fact     |     |     Store Dim     |
|-------------------|     |-------------------|     |-------------------|
| Product ID (PK)   |<----+ Sales ID (PK)     +---->| Store ID (PK)     |
| Product Name      | 1:N | Order Date        | 1:N | Store Name        |
| Category, Brand   |     | Customer ID (FK)  |     | Region, City      |
| Cost, Stock       |     | Product ID (FK)   |     | Store Type        |
+-------------------+     | Store ID (FK)     |     +-------------------+
                          | Quantity          |
                          | Unit Price        |
                          | Discount (%)      |
                          | TOTAL AMOUNT      |
                          +-------------------+
```

### Table Details & Field Descriptions
| Table Name | Primary / Foreign Key | Key Fields | Description |
| :--- | :--- | :--- | :--- |
| **Customer Dim** | `Customer ID` (PK) | Name, Age, Gender, City, State, Country, Loyalty Level | Demographic and tier classification of customers. |
| **Product Dim** | `Product ID` (PK) | Product Name, Category, Sub-Category, Brand, Cost, Stock | Product catalogue detailing pricing, inventory, and category. |
| **Store Dim** | `Store ID` (PK) | Store Name, Region, City, Store Type | Retail outlet properties (Flagship, Outlet, Online). |
| **Sales Fact** | `Sales ID` (PK)<br>`Customer ID` (FK)<br>`Product ID` (FK)<br>`Store ID` (FK) | Order Date, Quantity, Unit Price, Discount, Total Amount | Core transactional records capturing metrics and key relationships. |

---

## 🛠️ Data Cleaning & Transformation Rules

### 1. Formatting & Cleaning Text
Excel formulas were combined to strip non-printable characters, extra whitespaces, and enforce Title Case across all text fields:
```excel
=PROPER(TRIM(CLEAN(Cell_Reference)))
```

### 2. Handling Missing Values (Imputation Strategy)
- **Loyalty Level:** Replaced missing entries with `"UNKNOWN"`.
  ```excel
  =IF(ISBLANK(H2), "UNKNOWN", H2)
  ```
- **Numeric Fields Imputation:** Missing values in `Cost`, `Stock`, `Quantity`, `Unit Price`, and `Discount` were imputed using mean values over their respective ranges:
  ```excel
  Cost Imputation       : =IF(ISBLANK(F2), AVERAGE(F2:F103), F2)
  Stock Imputation      : =IF(ISBLANK(G2), AVERAGE(G2:G103), G2)
  Quantity Imputation   : =IF(ISBLANK(F2), AVERAGE(F2:F2001), F2)
  Unit Price Imputation : =IF(ISBLANK(G2), AVERAGE(G2:G2001), G2)
  Discount Imputation   : =IF(ISBLANK(H2), AVERAGE(H2:H2001), H2)
  ```

### 3. Business Calculation
Total transaction sales amount calculated as:
$$\text{Total Amount} = \text{Quantity} \times \text{Unit Price} \times (1 - \text{Discount})$$

### 4. Dynamic Data Lookup Implementation
`XLOOKUP` and `VLOOKUP` were used to populate descriptive fields into consolidated analytical views:
- **XLOOKUP Formulation:**
  - Customer Name: `=XLOOKUP(C2, Customer_Dim!A:A, Customer_Dim!B:B)`
  - Product Name: `=XLOOKUP(D2, Product_Dim!A:A, Product_Dim!B:B)`
  - Store Name: `=XLOOKUP(E2, Store_Dim!A:A, Store_Dim!B:B)`
- **VLOOKUP Alternative:**
  - Customer Name: `=VLOOKUP(C2, Customer_Dim!A:B, 2, FALSE)`

---

## 📈 Descriptive Statistics Overview

Summary metrics derived from **2,000 transactions**:

| Metric | Quantity (Units) | Unit Price ($) | Total Revenue ($) |
| :--- | :--- | :--- | :--- |
| **Mean (Average)** | 2.48 | $513.98 | — |
| **Median** | 2.00 | $522.43 | — |
| **Mode** | 2.00 | $871.34 | — |
| **Standard Deviation** | 1.11 | $282.09 | — |
| **Min / Max** | 1 / 4 | $20.09 / $999.89 | — |
| **Sum (Total)** | **4,953 Units** | **$1,027,966.23** | **$2,170,101.00** |

---

## 📊 Analytical Insights & Storytelling

### 1. 📖 Descriptive Analysis (What Happened?)
- **Revenue Distribution:** Total revenue reached **$2,170,101**.
- **Top Categories:** **Sports** ($542,010) and **Electronics** ($505,315) emerged as the primary revenue drivers, followed by Home ($422,651), Clothing ($377,543), and Beauty ($322,582).
- **Demographics:** Female buyers accounted for **57% ($1,229,719)** of overall sales, outperforming Male buyers at **43% ($940,382)**.

### 2. 🔍 Diagnostic Analysis (Why Did It Happen?)
- **Channel Performance:** **Online stores** ($914,963) and **Flagship locations** ($845,543) significantly outpaced **Outlet stores** ($409,594).
- **Driver:** Online and Flagship channels offer higher accessibility and product availability, whereas Outlet stores underperform due to localized demand and weaker promotional reach.

### 3. 🔮 Predictive Analysis (What Is Likely To Happen?)
- **Regional Trends (2023 – 2025):** Historical data shows a massive revenue surge in 2024 across all regions (e.g., North reached $337,185; South reached $332,830).
- **Forecast:** Although slight normalizations occurred in 2025, the **East and North regions** demonstrate stable, resilient demand, making them dependable growth drivers for future quarters.

### 4. 💡 Prescriptive Analysis (What Should We Do?)
- **Loyalty Program Optimization:** **Platinum (30%)** and **Gold (24%)** loyalty tiers account for over half of total units sold. Providing tailored promotional packages and priority rewards to these tiers will maximize Customer Lifetime Value (CLV).
- **Targeted Store Interventions:** Revamp pricing and discount strategies in low-performing **Outlet stores** and underpenetrated regions (**West**) to capture unrealized market share.

---

## 📁 Repository Structure
```text
.
├── README.md                           # Project Documentation
├── data/
│   ├── Raw_Data.xlsx                   # Uncleaned dataset
│   └── Transformed_Sales_Data.xlsx     # Cleaned dataset with Star Schema
└── dashboard/
    └── E-Commerce_Sales_Dashboard.xlsx # Interactive Pivot Tables & Dashboards
```

---

## 🚀 How to Use / Reproduce
1. **Clone the Repository:**
   ```bash
   git clone https://github.com/your-username/ecommerce-excel-capstone.git
   ```
2. **Open workbook in Microsoft Excel:** Launch `E-Commerce_Sales_Dashboard.xlsx`.
3. **Explore Data Modeling:** Inspect sheet relationships across `Customer_Dim`, `Product_Dim`, `Store_Dim`, and `Sales_Fact`.
4. **Interact with Dashboards:** Utilize Excel Slicers (Region, Year, Loyalty Level) to filter interactive pivot charts dynamically.
