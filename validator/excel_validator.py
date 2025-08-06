import pandas as pd
from datetime import datetime
import re

# Dictionary of valid functions and their corresponding subfunctions
functions_dict = {
    "General Management": [
        "CEO",
        "Corporate/Region/Country General Management",
        "Chief of Staff",
        "Executive Assistant"
    ],
    "Human Resources": [
        "CHRO & HR Management",
        "HR business Partner",
        "HR advisory",
        "Compensation & Benefits",
        "Health & Safety",
        "Diversity, Equity & Inclusion",
        "L&D Service/Campaing training",
        "L&D Analysis, instructional design and content lab",
        "L&D Platforms, Administration & Reporting",
        "L&D Professional skills training and services",
        "Personnel Administration & Payroll",
        "Industrial Relations and Labor legislation",
        "Talent Acquisition",
        "HR Services & Data management/Reporting",
        "Talent & Career Management",
        "Local Legally Mandatory Resources",
        "Other",
        "HR Mandatory Apprentices"
    ],
    "Finance": [
        "CFO & Finance Management",
        "Accounting & Consolidation",
        "FP&A",
        "Controlling",
        "Treasury",
        "Tax",
        "Procurement",
        "Real Estate & Facilities",
        "Finance Mandatory Apprentices"
    ],
    "Sales": [
        "CCO & Sales Management",
        "Commercial Marketing",
        "Strategic Sales",
        "Strategic Growth",
        "Sales Operations",
        "Portfolio & Partners &TPAs",
        "Presales",
        "Growth Office",
        "Account Manager",
        "Geo Hunter",
        "Sales Mandatory Apprentices"
    ],
    "Marketing and Communications": [
        "Group Marketing",
        "Group Communications",
        "Regional/Countries Marketing",
        "Regional/Countries Communications",
        "Group Internal Creative Agency",
        "Marcom Mandatory Apprentices"
    ],
    "Information Technology (CIO)": [
        "CIO & IT Management",
        "Infrastructure",
        "Workplace",
        "Software",
        "Governance",
        "CIO Mandatory Apprentices"
    ],
    "Technology Services (CTO)": [
        "CTO & Technology Management",
        "Software Architecture",
        "Reliability engineering",
        "Data Architecture",
        "AI/ML Engineering",
        "Product Management",
        "Project Management",
        "Software Development",
        "CTO Mandatory Apprentices"
    ],
    "Chief Information Security Officer (CISO)": [
        "CISO management",
        "Cibersecurity (Mission)",
        "Cybersecurity GRC (Governance, Risk and Compliance)",
        "Google Security",
        "Network Security",
        "Red Team",
        "Blue Team",
        "CISO Mandatory Apprentices"
    ],
    "Legal & Compliance": [
        "CLO & Legal Management",
        "Legal",
        "Sustainability",
        "Risks & Compliance",
        "Foundation",
        "Legal & Compliance Mandatory Apprentices"
    ],
    "Global Transformation": [
        "Transformation Management Office",
        "Organization and Procedures",
        "Business intelligence (BI)",
        "GEN AI",
        "Standardization and Corporate Information Systems",
        "Global Transformation Mandatory Apprentices"
    ],
    "GDSU (Global Digital Services Unit )": [
        "GDSU Management",
        "Digital Marketing",
        "Employee experience",
        "Product & Services",
        "AI & GenAI Services",
        "Advisory & Consulting",
        "Digital Partnerships Execution"
    ],
    "GSDE (Global Service Delivery Excellence)": [
        "GSDE Management",
        "Operations management",
        "Customer Operations & Performance",
        "Training & Quality Assurance",
        "Project Management Office (PMO)",
        "GSDE Mandatory Apprentices"
    ],
    "Digital Operations": [
        "Digital Operations management",
        "Digital Marketing",
        "Employee experience",
        "Product & Services",
        "AI & GenAI Services",
        "Advisory & Consulting",
        "Digital Operations Mandatory Apprentices"
    ],
    "Operations": [
        "Operations management & Operations Strategy",
        "Supervisors",
        "Coordinators",
        "Agents",
        "Quality, Certifications and Performance Management",
        "Workforce Management & Reporting",
        "Analytics & Data science",
        "Operations Mandatory Apprentices",
        'OPERATIONS'
    ]
}


def validate_excel_file(excel_file_path: str, output_txt_path: str = "validation_errors.txt"):
    """
    Validates an Excel file according to specified business rules and generates an error report.
    
    Args:
        excel_file_path (str): Path to the Excel file to validate
        output_txt_path (str): Path for the output error report file
    """
    
    # Expected headers in exact order
    expected_headers = [
        "COUNTRY", "YEAR", "MONTH", "EMPLOYEE_ID", "BIRTH DATE", "GENDER", "CITIZENSHIP", 
        "DISABILITY", "TRAINING", "COMPANY", "CORPORATION CODE", "CONTRACT TYPE", 
        "CONTRACT TIME", "AGREEMENT HOURS (CBA)", "CONTRACT /EMPLOYEE HOURS", "HIRING DATE", 
        "DISCHARGE DATE", "DISCHARGE CODE", "Temporary disability", "Maternity / Paternity leave", 
        "Labor union hours -", "Others Paid Abs", "Suspensions -", 
        "Others Unpaid (Justified reasons)", "Others Unpaid (Unjustified reasons)", 
        "WORK CENTRE", "COST CENTRE", "CATEGORY", "TOTAL COST_INVOICE\n(TOTAL LABOR COST)", 
        "GROSS SALARY", "VARIABLE / INCENTIVE REMUNERATION\n(Other local variable plans)", 
        "BONUS /SPORADIC PRIZES\n(MBO)", "OVERTIME PAY", "SOCIAL BENEFICTS", "DISMISSAL PAY", 
        "WAGE ARREARS", "SOCIAL CONTRIBUTION", "OTHERS COMPANY COSTS -", "PROVISION", 
        "OTHERS COMPANY COSTS 1. -", "OTHERS COMPANY COSTS 2. -", "OTHERS COMPANY COSTS 3. -", 
        "OTHERS COMPANY COSTS 4. -", "OTHERS COMPANY COSTS 5. -", "LIQUID SALARY -", 
        "FUNCTION", "SUBFUNCTION", "WORK MODALITY", "EMPLOYEE CLASIFICATION", "SCOPE", 
        "NAME", "SURNAME", "ANNUAL GROSS SALARY", "ANNUAL VARIABLE SALARY", "DIGITAL"
    ]
    
    # Valid values for various fields
    valid_countries = [27, 30, 41, 40, 31, 33, 25, 45, 43, 42, 90, 37, 29, 66, 70, 24, 34, 46, 22, 32, 20, 26, 64, 28, 10, 23, 21, 60]
    valid_genders = ['M', 'F']
    valid_citizenship = ['E', 'L']
    valid_disability = ['S', 'N']
    valid_training = ['S', 'N']
    valid_company = ['K', 'E', 'A', 'S']
    valid_contract_type = ['I', 'T', 'S']
    valid_contract_time = ['T', 'P']
    valid_discharge_codes = ['V', 'F', 'D', 'N']
    valid_work_modality = ['WORK FROM HOME', 'ON SITE', 'MIX']
    valid_employee_classification = ['INDIRECT STRUCTURE', 'DIRECT OPERATIONS', 'BUSINESS SUPPORT']
    valid_scope = ['GLOBAL', 'LOCAL']
    
    # Convert functions_dict keys and values to uppercase for case-insensitive comparison
    functions_dict_upper = {}
    for func, subfuncs in functions_dict.items():
        functions_dict_upper[func.upper()] = [sf.upper() for sf in subfuncs]
    
    # Cost columns that need to sum up to TOTAL COST_INVOICE
    cost_columns = [
        "GROSS SALARY", "VARIABLE / INCENTIVE REMUNERATION\n(Other local variable plans)",
        "BONUS /SPORADIC PRIZES\n(MBO)", "OVERTIME PAY", "SOCIAL BENEFICTS", "DISMISSAL PAY",
        "WAGE ARREARS", "SOCIAL CONTRIBUTION", "OTHERS COMPANY COSTS -", "PROVISION",
        "OTHERS COMPANY COSTS 1. -", "OTHERS COMPANY COSTS 2. -", "OTHERS COMPANY COSTS 3. -",
        "OTHERS COMPANY COSTS 4. -", "OTHERS COMPANY COSTS 5. -", "LIQUID SALARY -"
    ]
    
    errors = []
    
    try:
        # Read Excel file
        df = pd.read_excel(excel_file_path)
        
        # Clean all string columns - strip whitespace and convert to uppercase for case-insensitive comparison
        string_columns = ['GENDER', 'CITIZENSHIP', 'DISABILITY', 'TRAINING', 'COMPANY', 
                         'CONTRACT TYPE', 'CONTRACT TIME', 'DISCHARGE CODE', 'FUNCTION', 
                         'SUBFUNCTION', 'WORK MODALITY', 'EMPLOYEE CLASIFICATION', 'SCOPE', 
                         'NAME', 'SURNAME', 'DIGITAL']
        
        for col in string_columns:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip().str.upper()
        
        # Check headers order
        actual_headers = df.columns.tolist()
        if actual_headers != expected_headers:
            errors.append("HEADER ERROR: Headers do not match expected order or names")
            errors.append(f"Expected: {expected_headers}")
            errors.append(f"Actual: {actual_headers}")
            errors.append("=" * 80)
        
        # Get current date for year/month validation
        current_date = datetime.now()
        if current_date.month == 1:
            expected_year = current_date.year - 1
            expected_month = 12
        else:
            expected_year = current_date.year
            expected_month = current_date.month - 1
        
        # Validate each row
        for index, row in df.iterrows():
            row_num = index + 2  # +2 because Excel rows start at 1 and we have header
            
            # Country validation
            try:
                country_val = int(row['COUNTRY'])
                if country_val not in valid_countries:
                    errors.append(f"Row {row_num}, COUNTRY: '{country_val}' is not in valid country list")
            except (ValueError, TypeError):
                errors.append(f"Row {row_num}, COUNTRY: '{row['COUNTRY']}' is not a valid integer")
            
            # Check if all countries in the file are the same
            unique_countries = df['COUNTRY'].nunique()
            if unique_countries > 1:
                errors.append(f"COUNTRY: Multiple countries found in file. Only one country value should be present")
            
            # Year validation
            try:
                year_val = int(row['YEAR'])
                if year_val != expected_year:
                    errors.append(f"Row {row_num}, YEAR: '{year_val}' should be '{expected_year}' (previous month's year)")
            except (ValueError, TypeError):
                errors.append(f"Row {row_num}, YEAR: '{row['YEAR']}' is not a valid year format (YYYY)")
            
            # Month validation
            try:
                month_val = int(row['MONTH'])
                if month_val != expected_month:
                    errors.append(f"Row {row_num}, MONTH: '{month_val}' should be '{expected_month:02d}' (previous month)")
            except (ValueError, TypeError):
                errors.append(f"Row {row_num}, MONTH: '{row['MONTH']}' is not a valid month format (MM)")
            
            # Birth Date validation
            birth_date_str = str(row['BIRTH DATE'])
            if not re.match(r'^\d{8}$', birth_date_str):
                errors.append(f"Row {row_num}, BIRTH DATE: '{birth_date_str}' is not in YYYYMMDD format")
            else:
                try:
                    birth_date = datetime.strptime(birth_date_str, '%Y%m%d')
                    age = (current_date - birth_date).days / 365.25
                    if age > 75:
                        errors.append(f"Row {row_num}, BIRTH DATE: Employee is {age:.1f} years old (over 75)")
                    elif age < 15:
                        errors.append(f"Row {row_num}, BIRTH DATE: Employee is {age:.1f} years old (under 15)")
                except ValueError:
                    errors.append(f"Row {row_num}, BIRTH DATE: '{birth_date_str}' is not a valid date")
            
            # Gender validation
            if row['GENDER'] not in valid_genders:
                errors.append(f"Row {row_num}, GENDER: '{row['GENDER']}' must be 'M' or 'F'")
            
            # Citizenship validation
            if row['CITIZENSHIP'] not in valid_citizenship:
                errors.append(f"Row {row_num}, CITIZENSHIP: '{row['CITIZENSHIP']}' must be 'E' or 'L'")
            
            # Disability validation
            if row['DISABILITY'] not in valid_disability:
                errors.append(f"Row {row_num}, DISABILITY: '{row['DISABILITY']}' must be 'S' or 'N'")
            
            # Training validation
            if row['TRAINING'] not in valid_training:
                errors.append(f"Row {row_num}, TRAINING: '{row['TRAINING']}' must be 'S' or 'N'")
            
            # Company validation
            if row['COMPANY'] not in valid_company:
                errors.append(f"Row {row_num}, COMPANY: '{row['COMPANY']}' must be one of {valid_company}")
            
            # Contract Type validation
            if row['CONTRACT TYPE'] not in valid_contract_type:
                errors.append(f"Row {row_num}, CONTRACT TYPE: '{row['CONTRACT TYPE']}' must be one of {valid_contract_type}")
            
            # Contract Time validation
            if row['CONTRACT TIME'] not in valid_contract_time:
                errors.append(f"Row {row_num}, CONTRACT TIME: '{row['CONTRACT TIME']}' must be 'T' or 'P'")
            
            # Hiring Date validation
            hiring_date_str = str(row['HIRING DATE'])
            if not re.match(r'^\d{8}$', hiring_date_str):
                errors.append(f"Row {row_num}, HIRING DATE: '{hiring_date_str}' is not in YYYYMMDD format")
            else:
                try:
                    hiring_date = datetime.strptime(hiring_date_str, '%Y%m%d')
                    reporting_period = datetime(expected_year, expected_month, 1)
                    if hiring_date > reporting_period:
                        errors.append(f"Row {row_num}, HIRING DATE: '{hiring_date_str}' cannot be after reporting period")
                except ValueError:
                    errors.append(f"Row {row_num}, HIRING DATE: '{hiring_date_str}' is not a valid date")
            
            # DISCHARGE DATE validation
            end_date_str = str(row['DISCHARGE DATE'])
            if pd.notna(row['DISCHARGE DATE']) and end_date_str.lower() != 'nan':
                if not re.match(r'^\d{8}$', end_date_str):
                    errors.append(f"Row {row_num}, DISCHARGE DATE: '{end_date_str}' is not in YYYYMMDD format")
                else:
                    try:
                        end_date = datetime.strptime(end_date_str, '%Y%m%d')
                        hiring_date = datetime.strptime(hiring_date_str, '%Y%m%d')
                        if end_date < hiring_date:
                            errors.append(f"Row {row_num}, DISCHARGE DATE: '{end_date_str}' cannot be before hiring date")
                    except ValueError:
                        errors.append(f"Row {row_num}, DISCHARGE DATE: '{end_date_str}' is not a valid date")
            
            # Discharge Code validation
            discharge_code = row['DISCHARGE CODE']
            end_date_filled = pd.notna(row['DISCHARGE DATE']) and str(row['DISCHARGE DATE']).lower() != 'nan'
            discharge_filled = pd.notna(discharge_code) and str(discharge_code).lower() != 'nan'
            
            if discharge_filled and discharge_code not in valid_discharge_codes:
                errors.append(f"Row {row_num}, DISCHARGE CODE: '{discharge_code}' must be one of {valid_discharge_codes}")
            
            if end_date_filled and not discharge_filled:
                errors.append(f"Row {row_num}: DISCHARGE DATE is filled but DISCHARGE CODE is null")
            
            if discharge_filled and not end_date_filled:
                errors.append(f"Row {row_num}: DISCHARGE CODE is filled but DISCHARGE DATE is null")
            
            # Cost columns validation
            total_cost = 0
            for col in cost_columns:
                try:
                    cost_val = float(row[col]) if pd.notna(row[col]) else 0
                    total_cost += cost_val
                except (ValueError, TypeError):
                    errors.append(f"Row {row_num}, {col}: '{row[col]}' must be a number")
            
            # Total cost validation
            try:
                expected_total = float(row["TOTAL COST_INVOICE\n(TOTAL LABOR COST)"])
                if abs(total_cost - expected_total) > 0.01:  # Allow small floating point differences
                    errors.append(f"Row {row_num}, TOTAL COST_INVOICE: '{expected_total}' does not equal sum of cost columns '{total_cost:.2f}'")
            except (ValueError, TypeError):
                errors.append(f"Row {row_num}, TOTAL COST_INVOICE: '{row['TOTAL COST_INVOICE\\n(TOTAL LABOR COST)']}' must be a number")
            
            # Function validation
            function_val = row['FUNCTION']
            if function_val not in functions_dict_upper.keys():
                errors.append(f"Row {row_num}, FUNCTION: '{function_val}' is not a valid function")
            
            # Subfunction validation
            subfunction_val = row['SUBFUNCTION']
            if function_val in functions_dict_upper and subfunction_val not in functions_dict_upper[function_val]:
                errors.append(f"Row {row_num}, SUBFUNCTION: '{subfunction_val}' is not valid for function '{function_val}'")
            
            # Work Modality validation
            if row['WORK MODALITY'] not in valid_work_modality:
                errors.append(f"Row {row_num}, WORK MODALITY: '{row['WORK MODALITY']}' must be one of {valid_work_modality}")
            
            # Employee Classification validation
            employee_class = row['EMPLOYEE CLASIFICATION']
            if employee_class not in valid_employee_classification:
                errors.append(f"Row {row_num}, EMPLOYEE CLASIFICATION: '{employee_class}' must be one of {valid_employee_classification}")
            
            # Employee Classification vs Function validation
            if function_val in ['DIGITAL OPERATIONS', 'OPERATIONS']:
                if employee_class not in ['DIRECT OPERATIONS', 'BUSINESS SUPPORT']:
                    errors.append(f"Row {row_num}: Function '{function_val}' must have EMPLOYEE CLASIFICATION as 'Direct Operations' or 'Business Support'")
            else:
                if employee_class != 'INDIRECT STRUCTURE':
                    errors.append(f"Row {row_num}: Function '{function_val}' must have EMPLOYEE CLASIFICATION as 'Indirect structure'")
            
            # Scope validation
            scope_val = row['SCOPE']
            if scope_val not in valid_scope:
                errors.append(f"Row {row_num}, SCOPE: '{scope_val}' must be 'Global' or 'Local'")
            
            # Scope vs Function validation
            if function_val in ['GDSU (GLOBAL DIGITAL SERVICES UNIT )', 'GSDE (GLOBAL SERVICE DELIVERY EXCELLENCE)']:
                if scope_val != 'GLOBAL':
                    errors.append(f"Row {row_num}: Function '{function_val}' must have SCOPE as 'Global'")
            elif function_val in ['OPERATIONS', 'DIGITAL OPERATIONS']:
                if scope_val != 'LOCAL':
                    errors.append(f"Row {row_num}: Function '{function_val}' must have SCOPE as 'Local'")
            
            # Annual salary validations
            try:
                float(row['ANNUAL GROSS SALARY'])
            except (ValueError, TypeError):
                errors.append(f"Row {row_num}, ANNUAL GROSS SALARY: '{row['ANNUAL GROSS SALARY']}' must be a number")
            
            try:
                float(row['ANNUAL VARIABLE SALARY'])
            except (ValueError, TypeError):
                errors.append(f"Row {row_num}, ANNUAL VARIABLE SALARY: '{row['ANNUAL VARIABLE SALARY']}' must be a number")
            
            # Digital validation
            if function_val == 'DIGITAL OPERATIONS':
                if row['DIGITAL'] != 'Y':
                    errors.append(f"Row {row_num}: Function 'Digital Operations' must have DIGITAL value as 'Y'")
        
        # Write errors to file
        with open(output_txt_path, 'w', encoding='utf-8') as f:
            if errors:
                f.write("VALIDATION ERRORS FOUND:\n")
                f.write("=" * 80 + "\n\n")
                for error in errors:
                    f.write(error + "\n")
                f.write(f"\nTotal errors found: {len(errors)}")
            else:
                f.write("NO VALIDATION ERRORS FOUND!\n")
                f.write("All data validates successfully according to the specified rules.")
        
        print(f"Validation complete. Report saved to: {output_txt_path}")
        print(f"Total errors found: {len(errors)}")
        
    except Exception as e:
        error_msg = f"Error processing file: {str(e)}"
        with open(output_txt_path, 'w', encoding='utf-8') as f:
            f.write(error_msg)
        print(error_msg) 