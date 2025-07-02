# Automated Database Update for Label Printing

This project was developed for a logistics company to automate and accelerate the process of updating their label printing database. Previously, updating the database and preparing labels for shipments took several days and was prone to manual errors. With this solution, the process now takes just minutes, leveraging the power of Python, pandas, and Azure Web Apps.

## Problem Solved

- **Manual, slow, and error-prone process:** The company's staff had to manually process Excel files, apply complex business rules, and update the label printing database.
- **Delays in logistics:** Label generation for shipments could take days, slowing down the entire logistics chain.
- **Lack of automation:** No easy way to upload, process, and download updated database files.

## Solution

- **Web-based automation:** A Flask web app allows users to upload the current database and new load files.
- **Automated processing:** The app applies all business rules (barcode formatting, code splitting, quantity logic, row insertion, and more) using pandas.
- **Instant results:** The updated database is ready for download in seconds, not days.
- **Cloud deployment:** Hosted on Azure Web Apps for reliability, scalability, and easy access from anywhere.

## Features

- Upload the current database and new load files via a modern web interface.
- Automated data transformation and validation.
- Appends processed data to the database with correct column alignment and business logic.
- Download the updated database file for immediate use in label printing.
- Secure, scalable, and production-ready for logistics operations.

## Technologies Used

- **Python 3**
- **Flask**
- **pandas**
- **openpyxl**
- **Azure Web Apps**

## How to Deploy

1. Clone this repository:
   ```sh
   git clone https://github.com/saadhaniftaj/Automated-Database-update-for-label-printing.git
   cd Automated-Database-update-for-label-printing
   ```

2. Install dependencies:
   ```sh
   pip install -r requirements.txt
   ```

3. Run locally:
   ```sh
   python app.py
   ```
   Then open [http://localhost:8000](http://localhost:8000) in your browser.

4. Deploy to Azure Web Apps (see project comments for details).

## Impact

- Reduced label preparation time from days to minutes.
- Eliminated manual errors.
- Enabled the logistics team to focus on higher-value tasks.

---

**Developed for Nissi Distribution, ohio by saadhaniftaj.** 
