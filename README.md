# Konecta Data Validator

A Django web application for validating Excel files according to Konecta's business rules.

## Features

- Upload Excel files for validation
- Comprehensive data validation against business rules
- Downloadable validation reports
- Modern and user-friendly interface

## Setup

1. Install Python 3.8 or higher
2. Install the required packages:
   ```bash
   pip install -r requirements.txt
   ```
3. Place your company logo as `logo.jpg` in the `static` directory
4. Run the development server:
   ```bash
   python manage.py migrate
   python manage.py runserver
   ```
5. Access the application at `http://localhost:8000`

## Usage

1. Click "Choose Excel File" to select your Excel file
2. Click "Validate File" to start the validation process
3. If any errors are found, they will be displayed and available for download
4. Download the validation report for a detailed list of any issues found

## Business Rules

The validator checks for:
- Correct header order and names
- Valid country codes
- Date formats and ranges
- Employee information validity
- Cost calculations
- Function and classification relationships
- And many more business-specific rules

## Development

- Built with Django 5.0+
- Uses pandas for Excel processing
- Implements comprehensive error handling
- Modern UI with responsive design 