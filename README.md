# Customer-Information-and-Order-Management-in-Excel.
![20250520_1610_Customer Order Management_simple_compose_01jvpwz0ckec2a25p3kgg8ag0a](https://github.com/user-attachments/assets/ab01fc12-d09f-4e4d-9590-848c96cef195)

## Introduction 
Managing customer and order data across scattered spreadsheets can become a time-consuming routine for a sales support team in a bustling retail environment. Repetitive lookups, manual filtering, and endless scrolling often lead to errors and delays. To solve this, this project focuses on an Excel-based solution that automates the retrieval of customer information and their related orders through a centralized dashboard by simply entering a customer name in a dashboard interface. This approach reduces manual effort, enhances accuracy, and improves accessibility for business users who need to look up customer records and order history on demand. 


---


### Workbook Structure

<img width="951" alt="raw data" src="https://github.com/user-attachments/assets/2297f05d-648f-4ef7-9a29-ebd58666db10" />


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

Several entries were missing `Fax` numbers. To maintain consistency and clearly indicate unavailable information, these cells were filled with *Not Stated* using the **Find and Replace** tool. This approach ensures the dataset remains clean and avoids leaving blank fields that could disrupt analysis or presentation.


<img width="960" alt="image" src="https://github.com/user-attachments/assets/ba1c8dfa-e140-46aa-86f3-583aa9a3c118" />


---

## 3. Customer Dashboard Creation

To streamline access to individual customer records and their associated order history, a fully interactive **Customer Dashboard** was developed using Excel’s built-in tools and formulas.

<img width="951" alt="dashboard" src="https://github.com/user-attachments/assets/1ee10bea-8685-424b-9f79-d4b34794cbd6" />


### a. Customer Selection via Data Validation

I utilized the **Data Validation** feature from the **Data** tab to create a dropdown list for selecting customer names. This restricted the input to only valid, predefined names from the **Customer Info** sheet. By doing so, I prevented issues such as typos, partial names, or special characters, which could otherwise lead to lookup errors or incorrect data retrieval. This validation step not only enhanced data accuracy but also improved the reliability of the dashboard by ensuring only legitimate customer entries could be selected.

<img width="506" alt="validation" src="https://github.com/user-attachments/assets/050ce098-1124-4489-ba17-367f4cc28dca" />


### b. Populating Customer Information with XLOOKUP and VLOOKUP

Once a customer name is selected, Excel formulas retrieve the corresponding customer details automatically. I applied both **`XLOOKUP()`** and **`VLOOKUP()`** functions to fetch information such as:

* **Customer ID**
* **Contact Name**
* **Contact Title**
* **Address**
* **City**
* **Region**
* **Postal Code**
* **State**
* **Phone**
* **Fax**

<img width="521" alt="xlookup" src="https://github.com/user-attachments/assets/9b20cad7-952f-4293-991e-80f8c420c103" />


<img width="505" alt="vlookup" src="https://github.com/user-attachments/assets/fa45eb9c-c990-44d6-9aad-00f9b7f27cd4" />


These lookups are dynamically linked to the dropdown selection, ensuring that the dashboard updates instantly whenever a different customer is selected. This provided a seamless and user-friendly experience.

### c. Summary Metrics Using SUBTOTAL

To summarize key order metrics for each customer, I used the **`SUBTOTAL()`** function. This included:

* **Order Count**
* **Average Freight**
* **Last Order Date**

<img width="632" alt="subtotal" src="https://github.com/user-attachments/assets/2d51fb43-4a5f-495e-823b-3b26253ff7df" />


Unlike standard aggregation functions, `SUBTOTAL()` adjusts automatically based on filtered data making it ideal for dashboards that rely on interactivity.

### d. Order History Integration

The **Order History** section was brought in from the **Order Info** sheet. It lists all orders related to the selected customer. This section was connected using advanced filter logic and dynamic references so it could respond to changes in the customer selection.

### e. Automating Filters with Macros

Initially, the dashboard required manual refreshing of filters whenever a new customer was selected. To improve usability, I created a **macro-enabled button** that triggers an **Advanced Filter**. This automation instantly updates the order history and summary fields based on the current customer selection, eliminating the need for repetitive steps and improving the overall efficiency of the dashboard.

<img width="951" alt="populated dashboard" src="https://github.com/user-attachments/assets/0a944b79-b004-4141-bcdf-b9afa00265a5" />


---


### Step-by-Step Guide: Automating the Advanced Filter with a Macro

To automate refreshing the customer order history when a customer is selected, I created a macro tied to a button on the dashboard. Below are the steps I followed:


#### **Step 1: Enable Developer Tab**

If not already visible, go to:

* `File` → `Options` → `Customize Ribbon`
* Check the **Developer** checkbox to enable the tab.


#### **Step 2: Record the Macro**

1. Go to the **Developer** tab → click **Record Macro**.
2. In the pop-up:

   * **Macro Name:** `Advanced_Filter`or `FilterOrders`
   * **Store Macro In:** `This Workbook`
3. Click **OK** to start recording.
4. While recording, apply the **Advanced Filter** that links the selected customer to their corresponding order history from the **Order Info** sheet.
5. Stop the filter once satisfied.

<img width="960" alt="Record Macro" src="https://github.com/user-attachments/assets/6a8668e1-2db7-4fc3-9def-adf98524a48a" />


#### **Step 3: Stop Recording**

* Return to the Developer tab and click **Stop Recording** to save the macro.


#### **Step 4: Add and Assign Macro to a Button**

1. Still under the **Developer** tab, click **Insert** → choose **Button (Form Control)**.
2. Draw the button on the **Customer Dashboard** sheet.
3. In the dialog box that appears, choose the `FilterOrders` macro.
4. Rename the button to something intuitive like **“Load Data”** or **“Filter Orders”**.

<img width="960" alt="assign macro to button" src="https://github.com/user-attachments/assets/637db266-060f-4ba6-a443-817b025208f4" />


#### **Step 5: Use the Macro**

Clicking the macro button now triggers the **Advanced Filter**. This:

* Instantly updates the **Order History** to show only the selected customer’s records.
* Refreshes **Order Count**, **Average Freight**, and **Last Order** using `SUBTOTAL()`, making the dashboard fully dynamic.

---

### Outcome
Upon completion, this Excel-based tool serves as a lightweight, yet powerful customer lookup system suitable for small businesses or departments that handle customer interactions, support, or sales reporting. It enables staff to quickly access customer profiles and purchase history with minimal effort, supporting better customer service and operational decision-making.
