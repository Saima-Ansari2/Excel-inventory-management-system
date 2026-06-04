# Excel Inventory Management System

## Overview

This project is an Excel-based Inventory Management System developed using Microsoft Excel, VBA Macros, Data Validation, and Excel Formulas. The system automates purchase and sales entries, maintains stock records, and prevents invalid transactions through stock validation.

## Features

* Purchase Entry Management
* Sales Entry Management
* Automatic Stock Calculation
* VBA Macro Automation
* Auto Bill Number Generation
* Data Validation for User Inputs
* Stock Availability Validation
* Out of Stock Alert System
* Navigation Buttons using Hyperlinks
* Inventory Tracking Dashboard

## Functionalities

### Purchase Entry

* Records purchase transactions.
* Automatically updates stock quantity.
* Maintains purchase history.

### Sales Entry

* Records sales transactions.
* Deducts sold quantity from available stock.
* Updates inventory automatically.

### Stock Management

* Calculates available stock using Excel formulas.
* Displays current stock status for each item.
* Refresh functionality to update stock records.

### Stock Validation

* Prevents sales when sufficient stock is not available.
* Displays an "Out of Stock" alert message for invalid sales attempts.

## Technologies Used

* Microsoft Excel
* VBA Macros
* Data Validation
* SUMIF Formulas
* Hyperlinks
* Conditional Logic (IF Functions)

## Project Screenshots

### Home Screen

Displays entry forms and navigation controls.

### Purchase Entry

Used for recording inventory purchases.

### Sales Entry

Used for recording inventory sales.

### Stock Details

Shows available stock for all items.

### Out of Stock Alert

The system automatically validates inventory before processing a sale and displays an alert when stock is insufficient.

## How to Use

1. Open the `.xlsm` file.
2. Enable Macros when prompted.
3. Select Entry Type (Purchase or Sale).
4. Enter Bill Number, Date, Item, Quantity, and Rate.
5. Click the respective entry button.
6. Refresh stock records when required.
7. Check Stock Details for current inventory status.

## Note

This workbook contains VBA Macros. Please enable macros after opening the file to use all automation features.
