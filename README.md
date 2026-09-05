# Excel Customer Transactions & Financial Analysis

A comprehensive data analysis and data cleaning project built using **Microsoft Excel**. This project covers the complete workflow from raw data cleaning and transformation to advanced Excel functions, lookups, pivot tables, and conditional formatting.

---

## Project Overview

This project analyzes a customer transactions dataset to ensure data integrity and extract meaningful financial insights. It follows a structured set of business and data cleaning requirements.

---

## Key Tasks & Implementation

### 1. Data Cleaning & Preparation
* **Duplicate Removal:** Cleaned the dataset by removing duplicate entries based on `TransactionID`.
* **Missing Value Handling:** Filtered out and deleted rows with missing values in critical columns to maintain data quality.
* **Standardization:** Resolved capitalization and text consistency issues (e.g., standardizing text variants like "Savings" vs "savings").
* **Data Types:** Ensured all columns are formatted with correct data types (dates, text, currency).

### 2. Excel Functions & Advanced Operations
* **Lookup Functions:** Used Excel lookup formulas to retrieve `Branch` and `TransactionType` into Sheet1 based on `CustomerID`.
* **Date & Time Functions:** Extracted date components (Year, Month, Day) from `TransactionDate` for temporal analysis.
* **Logical Scenarios:** Applied conditional statements to classify transaction amounts (e.g., labeling amounts $> 4000$ as "High Value" and others as "Low Value").
* **Aggregation & Analysis:** 
  * Calculated total transaction amounts grouped by `AccountType`.
  * Computed specific sums for "Savings" accounts and "Current" accounts filtered by the year 2021.
  * Counted high-value transactions exceeding 3000.

### 3. Reporting & Visualization
* **Pivot Tables:** Built summary tables to analyze:
  * Total transaction amount by Account Type.
  * Top 5 customers by total transaction amount.
  * Customer distribution across different branches.
* **Conditional Formatting:** Applied visual highlights to flag transactions where the amount exceeds 3000 for quick risk/audit review.

---

## Project Structure

```text
├── Customer_Transactions_Cleaned.xlsx   # The final cleaned dataset and worked workbook
├── Assignment_Requirements.docx         # Project guidelines and tasks
└── README.md                            # Project documentation
