# 📊 Data Analysis & Visualization (Excel)

This repository showcases my proficiency in **data analysis, cleaning, pivoting, and visualization** using **Microsoft Excel**. It contains three distinct projects focusing on different datasets to demonstrate a wide range of analytical and reporting capabilities.

## 🌟 Project Overview

The projects utilize practical Excel features, including **Pivot Tables**, **Calculated Fields**, **VLOOKUP** and **Charting/Visualization**, to transform raw data into actionable business intelligence.

## 💾 Datasets and Analysis Focus

I worked with three diverse datasets:

| Project | Original Source (Implied) | Core Analysis Focus |
| :--- | :--- | :--- |
| **1. Global Bike Sales** | `Bike_Sales_Visualizations.xlsx` | Revenue/Profit trend analysis, geographical sales performance, and customer segmentation (Age Group, Gender). |
| **2. Retail Sales Performance** | `retail_sales_dataset.xlsx` | Sales decomposition by product category, gender, and customer generation. Calculation of commissions and simple transaction summaries. |
| **3. Atlantic Hurricanes** | `Biggest Atlantic Hurricanes -1.xlsx` | Historical storm impact analysis, comparing damage, wind speed, and Saffir-Simpson category. |

***

## 🛠️ Key Skills Demonstrated

Here are the Excel data analysis and visualization techniques I used across these projects:

* **Data Cleaning & Preparation:** Ensuring data consistency and integrity across multiple sheets.
* **Pivot Tables & Pivot Charts:** Creating dynamic summaries and visualizations for multi-dimensional analysis (e.g., Sales by Country and Product Category).
* **Calculations & Metrics:** Deriving key performance indicators (KPIs) such as **Profit, Revenue, Total Sales**, and **Commission**.
* **Time Series Analysis:** Tracking **Annual Revenue and Profit** trends over a multi-year period (2017-2021).
* **Segmentation Analysis:** Analyzing performance metrics segmented by **Age Group, Gender, Country, and Product Category**.
* **Data Visualization:** Designing clear and impactful charts.

***

## 📈 Key Analysis & Insights (Examples)

### Project 1: Global Bike Sales

Analysis of the bike sales data revealed critical business drivers across different customer demographics and geographical regions.

* **Annual Performance:** Analysis of `Revenue and Profit by Year` shows a clear upward trend:
    * **Total Revenue** over the 5-year period exceeded **$95 Million**.
    * **Annual Profit** increased substantially from **$4.06 Million (2017)** to **$12.98 Million (2021)**, indicating strong growth.
* **Customer Segmentation:** `Revenue by Age Group` shows that **Adults (35-64)** are the primary revenue drivers, accounting for over **$47 Million** in revenue.
* **Product vs. Country:** `Product Revenue by Country` highlights that **Bikes** are the largest revenue category across all countries, with the **United States** generating the highest total revenue overall.

### Project 2: Retail Sales Performance

This project focused on understanding consumer behavior and sales contribution across different groups.

* **Segmentation by Generation:** `pivot table by generation` reveals:
    * **Adults** have the highest contribution to overall sales (approx. **41.6%**).
    * **Beauty** and **Electronics** are the most popular categories among the **Adult** and **Senior** generations.
* **Segmentation by Gender:** `Pivot Table` indicates a relatively even split in total sales contribution between genders, with **Electronics** being the highest-grossing category overall (approx. **34.4%** of total sales).

### Project 3: Biggest Atlantic Hurricanes

The hurricane data provides a historical perspective on the scale of major storms.

* **Categorical Logic (for Wind Speed):**
    * The **`SWITCH`** function was used to map the storm's Saffir-Simpson **Category** (a number 1-5) to its descriptive **Max Wind Speed** range.
    * **Formula Example:** `=SWITCH([Category], 5, "157 and over", 4, "156", 3, "129", 2, "110", 1, "95", "N/A")`

* **Key Finding:** Analysis focused on correlating storm **Category** and **Max Wind Speed** with reported **Damage (USD Millions)** to identify the most financially destructive and powerful storms on record.

***

***

## 🚀 How to View the Work

To fully appreciate the analysis and visualizations:

1.  **Clone this repository** to your local machine.
2.  Navigate to the relevant project folders.
3.  Open the original Excel files to examine the formulas, pivot table structure, and chart designs.
