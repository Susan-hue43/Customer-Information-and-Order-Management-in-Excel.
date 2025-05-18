# Customer-Information-and-Order-Management-in-Excel.
## Introduction 
Managing customer and order data across scattered spreadsheets can become a time-consuming routine for a sales support team in a bustling retail environment. Repetitive lookups, manual filtering, and endless scrolling often led to errors and delays. To solve this, this project focuses on an Excel-based solution that automates the retrieval of customer information and their related orders through a centralized dashboard by simply entering a customer name in a dashboard interface. This approach reduces manual effort, enhances accuracy, and improves accessibility for business users who need to look up customer records and order history on demand. 


---


### Workbook Structure

The Excel workbook is structured into three key worksheets:

* **Customer Info:** This sheet contains a database of customer information, with **91** rows and **11** columns. Each row represents a unique customer, and the columns store various attributes such as customer name, contact details, and other relevant data.

* **Order Info:** This worksheet logs all orders placed by customers. It contains **830** rows and **8** columns. Each row represents a specific order, linked to a customer in the `"Customer Info"` sheet, and includes details like order ID, date, and items.

* **Customer Info Dashboard:** This is an interactive sheet designed to provide a user-friendly interface for quickly retrieving customer information and their order history. By entering a customer's name, the dashboard will dynamically display their details, along with a summary of their orders.
  
#### Key Features and Functionality:
* **Data Organization:** The `"Customer Info"` and `"Order Info"` sheets are structured to maintain data integrity and facilitate efficient data retrieval.

* **Dynamic Dashboard:** The `"Customer Info Dashboard"` uses Excel formulas (including VLOOKUP, XlOOKUP) to link the data across the worksheets. This allows for real-time updates of customer information and order details based on the customer name entered.

* **User-Friendly Interface:** The dashboard is designed to be intuitive, enabling users to quickly find the information they need.


---


## Objectives

* Enable seamless lookup of customer data using an intuitive interface.

* Automate data retrieval to eliminate repetitive manual searches.

* Provide a consolidated view of customer and order information in one place.

* Improve accuracy and speed of customer data access for reporting and support tasks.


---


## Data Description

### 1. Customer Info Table

| **Column Name** | **Description**                                                                 |
| --------------- | ------------------------------------------------------------------------------- |
| `Company Name`  | The name of the company or organization associated with the customer.           |
| `Customer ID`   | A unique identifier for each customer, used to link to order records.           |
| `Contact Name`  | The full name of the primary contact person for the customer.                   |
| `Contact Title` | The job title or position of the contact person within their company.           |
| `Address`       | The street address or mailing address of the customer.                          |
| `City`          | The city in which the customer is located.                                      |
| `Region`        | The geographical region (if applicable) associated with the customer's address. |
| `Postal Code`   | The postal or ZIP code of the customer’s address.                               |
| `State`         | The state or province where the customer resides.                               |
| `Phone`         | The customer's main contact telephone number.                                   |
| `Fax`           | The customer's fax number (if provided).                                        |


### 2. Order Info Table

| **Column Name** | **Description**                                                                 |
| --------------- | ------------------------------------------------------------------------------- |
| `OrderID`       | A unique identifier assigned to each order.                                     |
| `CustomerID`    | A foreign key linking the order to the corresponding customer in Customer Info. |
| `OrderDate`     | The date on which the order was placed.                                         |
| `RequiredDate`  | The date by which the customer expects the order to be delivered.               |
| `ShippedDate`   | The date on which the order was shipped.                                        |
| `ShipVia`       | The shipping method or carrier used for delivering the order.                   |
| `OrderAmount`   | The total monetary value of the order.                                          |
| `ShipName`      | The name of the recipient or business where the order is shipped.               |

---

## Data Transformation

To prepare the dataset for reliable automated extraction and dashboard integration, the following transformations were applied to the `Customer Info` worksheet:

### 1. Inconsistent Capitalization

Several text fields such as `Company Name`, `Contact Name`, `Contact Title`, `City`, and `State` contained inconsistent capitalization (e.g., “alfreds freddy” vs “Bonbay Apparel”). This was standardized to **proper case** (first letter uppercase, the rest lowercase) using Excel’s `PROPER()` function to enhance readability and ensure uniform lookup behavior.

**Example Transformation:**
* Before: `alfreds freddy`
* After: `Alfreds Freddy`

### 2. Missing Values

#### a. Region

Several entries had blank `Region` values despite having valid `State` and `City` data. A **lookup table** was created to map states to their respective regions using the `XLOOKUP()` function. The classification followed commonly accepted geographic groupings in India based on cultural, historical, and administrative divisions:

* **North**: Delhi, Punjab, Rajasthan
* **South**: Tamil Nadu, Telengana, Kerela
* **East**: Bihar, Jharkhand, Odisha
* **West**: Gujarat, Maharashtra
* **Northeast**: Assam, Meghalaya, Mizoram, Tripura

This filled all missing values while preserving consistency for regional analysis in the dashboard.

#### b. Fax

A number of entries lacked `Fax` numbers. These were left **intentionally blank**, as fax data was not critical to the analysis or dashboard functions. However, null values were visually marked (e.g., using conditional formatting) to alert users.

---
