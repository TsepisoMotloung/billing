import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import tkinter as tk
from tkinter import filedialog
from datetime import datetime
import re
import warnings
import xlrd
import openpyxl
from pathlib import Path
import matplotlib.colors as mcolors

# Suppress warnings
warnings.simplefilter(action='ignore', category=FutureWarning)


def select_folder():
    """Open a folder selection dialog and return the selected folder path."""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    folder_path = filedialog.askdirectory()
    return folder_path


def extract_month_year(filename):
    """Extract month and year from filename."""
    # Try to match patterns like "Billing Jan 2025.xls" or "Receipts_Jan_2025.xlsx"
    match = re.search(
        r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s*(\d{4})',
        filename, re.IGNORECASE)
    if match:
        month = match.group(1)
        year = match.group(2)
        return f"{month} {year}"
    return None


def find_matching_file(files, month_year):
    """Find a file that contains the given month_year in its name."""
    for file in files:
        if month_year.lower() in file.lower():
            return file
    return None


def load_excel_file(file_path):
    """Load an Excel file into a pandas DataFrame, handling different Excel formats."""
    try:
        # Try to determine if it's an .xls or .xlsx file
        if file_path.lower().endswith('.xls'):
            # Handle Excel 97-2003 (.xls) files
            df = pd.read_excel(file_path, engine='xlrd')
        else:
            # Handle newer Excel (.xlsx) files
            df = pd.read_excel(file_path, engine='openpyxl')

        return df
    except Exception as e:
        print(f"Error loading {file_path}: {str(e)}")
        return pd.DataFrame()  # Return empty DataFrame on error


def determine_policy_column(df, is_receipt=False):
    """Determine the column name that contains policy numbers."""
    if is_receipt:
        # For receipt files
        policy_cols = [
            'PolicyNo', 'Policy No', 'Policy_No', 'Policy Number',
            'PolicyNumber'
        ]
    else:
        # For billing files
        policy_cols = [
            'Policy_Number', 'Policy Number', 'PolicyNumber', 'Policy No',
            'Policy#'
        ]

    # Check for exact matches first
    for col in policy_cols:
        if col in df.columns:
            return col

    # If no exact match, look for partial matches
    for col in df.columns:
        if 'policy' in col.lower():
            return col

    # If still not found, return the first column as a fallback
    if len(df.columns) > 0:
        return df.columns[0]

    return None


def determine_amount_column(df, is_receipt=False):
    """Determine the column name that contains amount information."""
    if is_receipt:
        # For receipt files
        amount_cols = [
            'ReceiptAmount', 'Receipt Amount', 'Amount', 'Payment',
            'PaymentAmount'
        ]
    else:
        # For billing files
        amount_cols = [
            'Premium', 'Premium Amount', 'Amount', 'Bill Amount',
            'BillingAmount'
        ]

    # Check for exact matches first
    for col in amount_cols:
        if col in df.columns:
            return col

    # If no exact match, look for partial matches
    for col in df.columns:
        if any(term in col.lower() for term in
               ['amount', 'premium', 'payment', 'receipt', 'bill']):
            return col

    # If still not found, return None
    return None


def determine_name_columns(df, is_receipt=False):
    """Determine the columns that contain client name information."""
    name_info = {}

    if is_receipt:
        # For receipt files
        if 'ClientName' in df.columns:
            name_info['full_name'] = 'ClientName'
        elif 'Firstname' in df.columns and 'Surname' in df.columns:
            name_info['first_name'] = 'Firstname'
            name_info['last_name'] = 'Surname'
        else:
            # Look for any name-related columns
            for col in df.columns:
                if 'name' in col.lower() or 'client' in col.lower(
                ) or 'customer' in col.lower():
                    name_info['full_name'] = col
                    break
    else:
        # For billing files
        if 'Surname' in df.columns and 'First Name' in df.columns:
            name_info['first_name'] = 'First Name'
            name_info['last_name'] = 'Surname'
        elif 'Surname' in df.columns and 'Initials' in df.columns:
            name_info['first_name'] = 'Initials'
            name_info['last_name'] = 'Surname'
        else:
            # Look for any name-related columns
            for col in df.columns:
                if 'name' in col.lower() or 'client' in col.lower(
                ) or 'customer' in col.lower():
                    name_info['full_name'] = col
                    break

    return name_info


def determine_product_column(df):
    """Determine the column that contains product information."""
    # The product information is in a column named "Client name"
    if "Client name" in df.columns:
        return "Client name"

    # Check for similar column names
    for col in df.columns:
        if "client name" in col.lower() or "clientname" in col.lower():
            return col

    # Check for other potential product columns as fallback
    product_cols = ["Product", "Plan", "Policy Type", "Plan Type"]
    for col in product_cols:
        if col in df.columns:
            return col

    return None


def get_client_name(row, name_info):
    """Extract client name from a row based on name column information."""
    if 'full_name' in name_info and name_info[
            'full_name'] in row and not pd.isna(row[name_info['full_name']]):
        return row[name_info['full_name']]
    elif 'first_name' in name_info and 'last_name' in name_info:
        first_name = row.get(name_info['first_name'], '')
        last_name = row.get(name_info['last_name'], '')

        if pd.isna(first_name):
            first_name = ''
        if pd.isna(last_name):
            last_name = ''

        if first_name or last_name:
            return f"{last_name}, {first_name}".strip(', ')

    return "Unknown"


def get_product_info(row, product_col):
    """Extract product information from a row."""
    if product_col and product_col in row and not pd.isna(row[product_col]):
        return row[product_col]
    return ""


def analyze_payments(billing_df, receipt_df, month_year, policy_history=None):
    """
    Compare billing and receipt data for a specific month/year.

    Args:
        billing_df: DataFrame containing billing data
        receipt_df: DataFrame containing receipt data
        month_year: String representing the month and year
        policy_history: Dictionary to track policy payment history across months

    Returns:
        results_df: DataFrame with analysis results
        summary: Dictionary with summary statistics
        policy_history: Updated policy history dictionary
    """
    if policy_history is None:
        policy_history = {}

    # Determine key columns in each DataFrame
    billing_policy_col = determine_policy_column(billing_df, is_receipt=False)
    receipt_policy_col = determine_policy_column(receipt_df, is_receipt=True)

    billing_amount_col = determine_amount_column(billing_df, is_receipt=False)
    receipt_amount_col = determine_amount_column(receipt_df, is_receipt=True)

    billing_name_info = determine_name_columns(billing_df, is_receipt=False)
    receipt_name_info = determine_name_columns(receipt_df, is_receipt=True)

    # Find product column (which is named "Client name")
    billing_product_col = determine_product_column(billing_df)
    receipt_product_col = determine_product_column(receipt_df)

    print(
        f"Billing columns identified: Policy={billing_policy_col}, Amount={billing_amount_col}, Product={billing_product_col}"
    )
    print(
        f"Receipt columns identified: Policy={receipt_policy_col}, Amount={receipt_amount_col}, Product={receipt_product_col}"
    )

    # Initialize results DataFrame
    results = []

    # Process billing records
    processed_receipts = set()

    # Initialize counters for summary
    correct_count = 0
    over_count = 0
    under_count = 0
    not_paid_count = 0
    not_billed_count = 0

    # Convert policy numbers to strings for comparison
    if billing_policy_col and billing_policy_col in billing_df.columns:
        billing_df[billing_policy_col] = billing_df[billing_policy_col].astype(
            str)
    if receipt_policy_col and receipt_policy_col in receipt_df.columns:
        receipt_df[receipt_policy_col] = receipt_df[receipt_policy_col].astype(
            str)

    # Process billing data
    for _, billing_row in billing_df.iterrows():
        if billing_policy_col not in billing_row or pd.isna(
                billing_row[billing_policy_col]):
            continue

        policy_number = str(billing_row[billing_policy_col]).strip()
        if not policy_number:
            continue

        # Get billed amount
        billed_amount = 0
        if billing_amount_col and billing_amount_col in billing_row:
            if pd.notna(billing_row[billing_amount_col]
                        ) and billing_row[billing_amount_col] != '':
                try:
                    billed_amount = float(billing_row[billing_amount_col])
                except (ValueError, TypeError):
                    billed_amount = 0

        # Initialize paid amount
        paid_amount = 0

        # Find matching receipt
        matching_receipt = receipt_df[receipt_df[receipt_policy_col] ==
                                      policy_number]

        if not matching_receipt.empty:
            # Get paid amount from matching receipt
            receipt_row = matching_receipt.iloc[0]
            if receipt_amount_col and receipt_amount_col in receipt_row:
                if pd.notna(receipt_row[receipt_amount_col]
                            ) and receipt_row[receipt_amount_col] != '':
                    try:
                        paid_amount = float(receipt_row[receipt_amount_col])
                    except (ValueError, TypeError):
                        paid_amount = 0

            # Mark as processed
            processed_receipts.add(policy_number)

        # Calculate difference and determine status
        difference = paid_amount - billed_amount

        if paid_amount == 0 and billed_amount > 0:
            status = "Not Paid"
            not_paid_count += 1
        elif abs(paid_amount - billed_amount
                 ) < 0.01:  # Using small epsilon for float comparison
            status = "Correct Payment"
            correct_count += 1
        elif paid_amount > billed_amount:
            status = "Overpaid"
            over_count += 1
        elif paid_amount < billed_amount and paid_amount > 0:
            status = "Underpaid"
            under_count += 1
        else:
            status = "Unknown"

        # Get client name and product
        client_name = get_client_name(billing_row, billing_name_info)
        product_name = get_product_info(billing_row, billing_product_col)

        # Add to results
        results.append({
            'Month_Year': month_year,
            'Policy_Number': policy_number,
            'Client_Name': client_name,
            'Product': product_name,
            'Billed_Amount': billed_amount,
            'Paid_Amount': paid_amount,
            'Difference': difference,
            'Status': status
        })

        # Track policy history
        if policy_number not in policy_history:
            policy_history[policy_number] = {
                'count': 1,
                'status': status,
                'consistent': True,
                'client_name': client_name,
                'product': product_name,
                f'{month_year}_billed': billed_amount,
                f'{month_year}_paid': paid_amount
            }
        else:
            # Update history
            history = policy_history[policy_number]
            history['count'] += 1

            # Check consistency
            if history['status'] != status and status != "Correct Payment":
                history['consistent'] = False

            # Update status (if not correct payment)
            if status != "Correct Payment":
                history['status'] = status

            # Add this month's data
            history[f'{month_year}_billed'] = billed_amount
            history[f'{month_year}_paid'] = paid_amount

            # Update client name and product if they were empty
            if not history.get('client_name'):
                history['client_name'] = client_name
            if not history.get('product') and product_name:
                history['product'] = product_name

    # Process receipts with no corresponding billing
    for _, receipt_row in receipt_df.iterrows():
        if receipt_policy_col not in receipt_row or pd.isna(
                receipt_row[receipt_policy_col]):
            continue

        policy_number = str(receipt_row[receipt_policy_col]).strip()
        if not policy_number or policy_number in processed_receipts:
            continue

        # Get paid amount
        paid_amount = 0
        if receipt_amount_col and receipt_amount_col in receipt_row:
            if pd.notna(receipt_row[receipt_amount_col]
                        ) and receipt_row[receipt_amount_col] != '':
                try:
                    paid_amount = float(receipt_row[receipt_amount_col])
                except (ValueError, TypeError):
                    paid_amount = 0

        # Only process if amount is greater than 0
        if paid_amount > 0:
            client_name = get_client_name(receipt_row, receipt_name_info)
            product_name = get_product_info(receipt_row, receipt_product_col)

            # Add to results
            results.append({
                'Month_Year': month_year,
                'Policy_Number': policy_number,
                'Client_Name': client_name,
                'Product': product_name,
                'Billed_Amount': 0,
                'Paid_Amount': paid_amount,
                'Difference': paid_amount,
                'Status': "Not Billed"
            })

            not_billed_count += 1

            # Track policy history
            if policy_number not in policy_history:
                policy_history[policy_number] = {
                    'count': 1,
                    'status': "Not Billed",
                    'consistent': True,
                    'client_name': client_name,
                    'product': product_name,
                    f'{month_year}_billed': 0,
                    f'{month_year}_paid': paid_amount
                }
            else:
                # Update history
                history = policy_history[policy_number]
                history['count'] += 1

                # Check consistency
                if history['status'] != "Not Billed":
                    history['consistent'] = False

                # Update status
                history['status'] = "Not Billed"

                # Add this month's data
                history[f'{month_year}_billed'] = 0
                history[f'{month_year}_paid'] = paid_amount

                # Update client name and product if they were empty
                if not history.get('client_name'):
                    history['client_name'] = client_name
                if not history.get('product') and product_name:
                    history['product'] = product_name

    # Create results DataFrame
    results_df = pd.DataFrame(results)

    # Calculate summary statistics
    total_policies = len(results)

    summary = {
        'Correct Payment': correct_count,
        'Overpaid': over_count,
        'Underpaid': under_count,
        'Not Paid': not_paid_count,
        'Not Billed': not_billed_count,
        'Total': total_policies
    }

    return results_df, summary, policy_history


def apply_consistency_status(results_df, policy_history):
    """Apply consistency status to the results DataFrame based on payment history."""
    consistency = []

    for _, row in results_df.iterrows():
        policy = row['Policy_Number']
        status = row['Status']

        if policy in policy_history:
            history = policy_history[policy]

            if history['count'] > 1 and history[
                    'consistent'] and status != "Correct Payment":
                consistency.append(f"Consistent {status}")
            elif status != "Correct Payment":
                consistency.append("Inconsistent")
            else:
                consistency.append("")
        else:
            consistency.append("")

    results_df['Consistency'] = consistency
    return results_df


def create_monthly_chart(summary, month_year, output_folder):
    """Create a pie chart for monthly payment summary."""
    labels = []
    sizes = []
    colors = ['#C6EF79', '#FFEB9C', '#FFC7CE', '#C00000', '#9BC2E6']

    for i, category in enumerate(
        ['Correct Payment', 'Overpaid', 'Underpaid', 'Not Paid',
         'Not Billed']):
        if summary[category] > 0:
            labels.append(f"{category} ({summary[category]})")
            sizes.append(summary[category])

    # Create pie chart
    plt.figure(figsize=(10, 6))
    plt.pie(sizes,
            labels=labels,
            colors=colors[:len(sizes)],
            autopct='%1.1f%%',
            startangle=140)
    plt.axis(
        'equal')  # Equal aspect ratio ensures that pie is drawn as a circle
    plt.title(f'Payment Status Analysis for {month_year}')

    # Save chart
    chart_path = os.path.join(output_folder,
                              f'chart_{month_year.replace(" ", "_")}.png')
    plt.savefig(chart_path)
    plt.close()

    return chart_path


def create_policy_matrix(policy_history, processed_months, output_folder):
    """Create a matrix showing policy performance across months."""
    # Sort months chronologically
    month_order = {
        'Jan': 1,
        'Feb': 2,
        'Mar': 3,
        'Apr': 4,
        'May': 5,
        'Jun': 6,
        'Jul': 7,
        'Aug': 8,
        'Sep': 9,
        'Oct': 10,
        'Nov': 11,
        'Dec': 12
    }

    def month_year_key(month_year):
        parts = month_year.split()
        month = parts[0]
        year = int(parts[1])
        return year * 100 + month_order.get(month, 0)

    sorted_months = sorted(processed_months, key=month_year_key)

    # Create DataFrame for matrix
    matrix_data = []

    for policy, history in policy_history.items():
        row_data = {
            'Policy_Number': policy,
            'Client_Name': history.get('client_name', ''),
            'Product': history.get('product', '')
        }

        # Add data for each month
        for month in sorted_months:
            billed_key = f'{month}_billed'
            paid_key = f'{month}_paid'

            if billed_key in history:
                row_data[f'{month}_Billed'] = history[billed_key]
            else:
                row_data[f'{month}_Billed'] = np.nan

            if paid_key in history:
                row_data[f'{month}_Paid'] = history[paid_key]
            else:
                row_data[f'{month}_Paid'] = np.nan

        matrix_data.append(row_data)

    matrix_df = pd.DataFrame(matrix_data)

    # Save matrix to CSV
    matrix_csv_path = os.path.join(output_folder, 'policy_matrix.csv')
    matrix_df.to_csv(matrix_csv_path, index=False)

    # Create a heatmap visualization of payment status
    create_payment_heatmap(matrix_df, sorted_months, output_folder)

    return matrix_csv_path


def create_payment_heatmap(matrix_df, months, output_folder):
    """Create a heatmap visualization of payment status across months."""
    # Prepare data for heatmap
    policies = matrix_df['Policy_Number'].tolist()

    # Calculate payment status for each policy/month
    status_data = []

    for _, row in matrix_df.iterrows():
        status_row = []

        for month in months:
            billed = row.get(f'{month}_Billed', np.nan)
            paid = row.get(f'{month}_Paid', np.nan)

            if pd.isna(billed) or pd.isna(paid):
                status_row.append(np.nan)  # No data
            elif billed == 0 and paid > 0:
                status_row.append(4)  # Not Billed
            elif paid == 0 and billed > 0:
                status_row.append(3)  # Not Paid
            elif abs(paid - billed) < 0.01:
                status_row.append(0)  # Correct Payment
            elif paid > billed:
                status_row.append(1)  # Overpaid
            elif paid < billed:
                status_row.append(2)  # Underpaid
            else:
                status_row.append(np.nan)

        status_data.append(status_row)

    # Create heatmap
    plt.figure(figsize=(max(12,
                            len(months) * 1.5), max(8,
                                                    len(policies) * 0.3)))

    # Create a custom colormap
    cmap = mcolors.ListedColormap(
        ['#C6EF79', '#FFEB9C', '#FFC7CE', '#C00000', '#9BC2E6'])

    # Create the heatmap
    ax = sns.heatmap(status_data,
                     cmap=cmap,
                     linewidths=0.5,
                     linecolor='gray',
                     cbar=False,
                     mask=np.isnan(status_data))

    # Set the labels
    ax.set_xticks(np.arange(len(months)) + 0.5)
    ax.set_xticklabels(months, rotation=45, ha='right')
    ax.set_yticks(np.arange(len(policies)) + 0.5)
    ax.set_yticklabels(policies)

    plt.title('Policy Payment Status by Month')

    # Add legend
    from matplotlib.patches import Patch
    legend_elements = [
        Patch(facecolor='#C6EF79', label='Correct Payment'),
        Patch(facecolor='#FFEB9C', label='Overpaid'),
        Patch(facecolor='#FFC7CE', label='Underpaid'),
        Patch(facecolor='#C00000', label='Not Paid'),
        Patch(facecolor='#9BC2E6', label='Not Billed')
    ]
    plt.legend(handles=legend_elements,
               loc='upper left',
               bbox_to_anchor=(1, 1))

    plt.tight_layout()

    # Save the heatmap
    heatmap_path = os.path.join(output_folder, 'payment_status_heatmap.png')
    plt.savefig(heatmap_path, dpi=300, bbox_inches='tight')
    plt.close()

    return heatmap_path


def create_summary_report(all_results, summary_by_month, policy_history,
                          output_folder):
    """Create a summary report with overall statistics."""
    # Create overall summary
    total_policies = len(policy_history)
    correct_count = sum(summary['Correct Payment']
                        for summary in summary_by_month.values())
    over_count = sum(summary['Overpaid']
                     for summary in summary_by_month.values())
    under_count = sum(summary['Underpaid']
                      for summary in summary_by_month.values())
    not_paid_count = sum(summary['Not Paid']
                         for summary in summary_by_month.values())
    not_billed_count = sum(summary['Not Billed']
                           for summary in summary_by_month.values())

    # Count consistent patterns
    over_consistent = 0
    under_consistent = 0

    for policy, history in policy_history.items():
        if history['count'] > 1 and history['consistent']:
            if history['status'] == 'Overpaid':
                over_consistent += 1
            elif history['status'] == 'Underpaid':
                under_consistent += 1

    # Create summary DataFrame
    summary_data = [{
        'Category': '1. Pay correctly as per bill',
        'Count': correct_count
    }, {
        'Category': '2. Pay over the billed amount',
        'Count': over_count
    }, {
        'Category': '3. Pay over the billed amount consistently',
        'Count': over_consistent
    }, {
        'Category': '4. Pay below the billed amount',
        'Count': under_count
    }, {
        'Category': '5. Pay below the billed amount consistently',
        'Count': under_consistent
    }, {
        'Category': '6. Billed not paid',
        'Count': not_paid_count
    }, {
        'Category': '7. Paid not billed',
        'Count': not_billed_count
    }, {
        'Category': 'Total Policies',
        'Count': total_policies
    }]

    summary_df = pd.DataFrame(summary_data)

    # Calculate percentages
    if total_policies > 0:
        summary_df['Percentage'] = summary_df['Count'] / total_policies
    else:
        summary_df['Percentage'] = 0

    # Save summary to CSV
    summary_csv_path = os.path.join(output_folder, 'summary.csv')
    summary_df.to_csv(summary_csv_path, index=False)

    # Create summary chart
    plt.figure(figsize=(10, 6))

    # Exclude the 'Total Policies' row and any zero-count categories
    chart_data = summary_df[summary_df['Category'] != 'Total Policies']
    chart_data = chart_data[chart_data['Count'] > 0]

    colors = [
        '#C6EF79', '#FFEB9C', '#FFEB9C', '#FFC7CE', '#FFC7CE', '#C00000',
        '#9BC2E6'
    ]

    plt.pie(chart_data['Count'],
            labels=chart_data['Category'],
            autopct='%1.1f%%',
            colors=colors[:len(chart_data)])
    plt.axis('equal')
    plt.title('Overall Payment Analysis')

    # Save chart
    chart_path = os.path.join(output_folder, 'summary_chart.png')
    plt.savefig(chart_path)
    plt.close()

    return summary_csv_path, chart_path


def create_html_report(output_folder, processed_months, summary_by_month):
    """Create an HTML report with embedded charts and links to CSV files."""
    html_content = """
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Billing and Receipt Analysis Report</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 20px; }
            h1, h2, h3 { color: #2c3e50; }
            .summary { margin-bottom: 30px; }
            .month-section { margin-bottom: 40px; border-bottom: 1px solid #eee; padding-bottom: 20px; }
            .chart-container { text-align: center; margin: 20px 0; }
            .chart-container img { max-width: 100%; height: auto; }
            table { border-collapse: collapse; width: 100%; margin: 20px 0; }
            th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
            th { background-color: #f2f2f2; }
            tr:nth-child(even) { background-color: #f9f9f9; }
            .download-link { display: inline-block; margin: 10px 0; padding: 8px 15px; 
                             background-color: #3498db; color: white; text-decoration: none; 
                             border-radius: 4px; }
            .download-link:hover { background-color: #2980b9; }
        </style>
    </head>
    <body>
        <h1>Billing and Receipt Analysis Report</h1>
        <div class="summary">
            <h2>Overall Summary</h2>
            <div class="chart-container">
                <img src="summary_chart.png" alt="Summary Chart">
            </div>
            <a class="download-link" href="summary.csv" download>Download Summary CSV</a>
            <a class="download-link" href="all_results.csv" download>Download All Results CSV</a>
            <a class="download-link" href="policy_matrix.csv" download>Download Policy Matrix CSV</a>
        </div>

        <h2>Policy Payment Status Heatmap</h2>
        <div class="chart-container">
            <img src="payment_status_heatmap.png" alt="Payment Status Heatmap">
        </div>

        <h2>Monthly Analysis</h2>
    """

    # Sort months chronologically
    month_order = {
        'Jan': 1,
        'Feb': 2,
        'Mar': 3,
        'Apr': 4,
        'May': 5,
        'Jun': 6,
        'Jul': 7,
        'Aug': 8,
        'Sep': 9,
        'Oct': 10,
        'Nov': 11,
        'Dec': 12
    }

    def month_year_key(month_year):
        parts = month_year.split()
        month = parts[0]
        year = int(parts[1])
        return year * 100 + month_order.get(month, 0)

    sorted_months = sorted(processed_months, key=month_year_key)

    # Add section for each month
    for month in sorted_months:
        safe_month = month.replace(" ", "_")
        summary = summary_by_month.get(month, {})

        html_content += f"""
        <div class="month-section">
            <h3>{month}</h3>
            <div class="chart-container">
                <img src="chart_{safe_month}.png" alt="Chart for {month}">
            </div>

            <h4>Summary</h4>
            <table>
                <tr>
                    <th>Category</th>
                    <th>Count</th>
                    <th>Percentage</th>
                </tr>
        """

        # Add rows for each category
        total = summary.get('Total', 0)
        if total > 0:
            for category in [
                    'Correct Payment', 'Overpaid', 'Underpaid', 'Not Paid',
                    'Not Billed'
            ]:
                count = summary.get(category, 0)
                percentage = (count / total) * 100 if total > 0 else 0

                html_content += f"""
                <tr>
                    <td>{category}</td>
                    <td>{count}</td>
                    <td>{percentage:.1f}%</td>
                </tr>
                """

        html_content += f"""
            </table>
            <a class="download-link" href="analysis_{safe_month}.csv" download>Download {month} CSV</a>
        </div>
        """

    # Close HTML
    html_content += """
    </body>
    </html>
    """

    # Write HTML file
    with open(os.path.join(output_folder, 'report.html'), 'w') as f:
        f.write(html_content)


def main():
    print("Billing and Receipt Analysis Tool")
    print("=================================")

    # Select folders
    print("\nPlease select the folder containing billing files:")
    billing_folder = select_folder()
    if not billing_folder:
        print("No billing folder selected. Exiting.")
        return

    print("\nPlease select the folder containing receipt files:")
    receipts_folder = select_folder()
    if not receipts_folder:
        print("No receipts folder selected. Exiting.")
        return

    # Create output folder for results
    output_folder = os.path.join(os.path.dirname(billing_folder),
                                 "Analysis_Results")
    os.makedirs(output_folder, exist_ok=True)

    print(f"\nAnalysis results will be saved to: {output_folder}")

    # Get list of Excel files
    billing_files = [
        f for f in os.listdir(billing_folder)
        if f.lower().endswith(('.xls', '.xlsx'))
    ]
    receipt_files = [
        f for f in os.listdir(receipts_folder)
        if f.lower().endswith(('.xls', '.xlsx'))
    ]

    if not billing_files:
        print("No Excel files found in the billing folder.")
        return

    if not receipt_files:
        print("No Excel files found in the receipts folder.")
        return

    print(
        f"\nFound {len(billing_files)} billing files and {len(receipt_files)} receipt files."
    )

    # Initialize variables to store results
    all_results = []
    summary_by_month = {}
    policy_history = {}
    processed_months = set()

    # Process each billing file
    for billing_file in billing_files:
        # Extract month/year from filename
        month_year = extract_month_year(billing_file)
        if not month_year:
            print(
                f"Could not extract month/year from {billing_file}. Skipping.")
            continue

        processed_months.add(month_year)

        # Find corresponding receipt file
        matching_receipt_file = find_matching_file(receipt_files, month_year)
        if not matching_receipt_file:
            print(
                f"No matching receipt file found for {month_year}. Skipping.")
            continue

        print(f"\nProcessing {month_year}:")
        print(f"  Billing file: {billing_file}")
        print(f"  Receipt file: {matching_receipt_file}")

        # Load data
        billing_df = load_excel_file(os.path.join(billing_folder,
                                                  billing_file))
        receipt_df = load_excel_file(
            os.path.join(receipts_folder, matching_receipt_file))

        if billing_df.empty or receipt_df.empty:
            print(f"Error loading data for {month_year}. Skipping.")
            continue

        # Analyze payments
        results_df, summary, policy_history = analyze_payments(
            billing_df, receipt_df, month_year, policy_history)

        # Apply consistency status
        results_df = apply_consistency_status(results_df, policy_history)

        # Store results and summary
        all_results.append(results_df)
        summary_by_month[month_year] = summary

        # Save monthly results to CSV
        month_csv_path = os.path.join(
            output_folder, f'analysis_{month_year.replace(" ", "_")}.csv')
        results_df.to_csv(month_csv_path, index=False)

        # Create chart for this month
        chart_path = create_monthly_chart(summary, month_year, output_folder)

        print(f"  Results saved to: {month_csv_path}")
        print(f"  Chart saved to: {chart_path}")
        print(f"  Summary: {summary}")

    # Combine all results
    if all_results:
        combined_results = pd.concat(all_results, ignore_index=True)
        combined_csv_path = os.path.join(output_folder, 'all_results.csv')
        combined_results.to_csv(combined_csv_path, index=False)
        print(f"\nCombined results saved to: {combined_csv_path}")

        # Create policy matrix
        matrix_path = create_policy_matrix(policy_history, processed_months,
                                           output_folder)
        print(f"Policy matrix saved to: {matrix_path}")

        # Create summary report
        summary_path, summary_chart_path = create_summary_report(
            combined_results, summary_by_month, policy_history, output_folder)
        print(f"Summary report saved to: {summary_path}")
        print(f"Summary chart saved to: {summary_chart_path}")

        print("\nAnalysis complete!")

        # Create HTML report with embedded charts
        create_html_report(output_folder, processed_months, summary_by_month)
        print(
            f"HTML report saved to: {os.path.join(output_folder, 'report.html')}"
        )
    else:
        print("No results were generated. Please check input data.")


if __name__ == "__main__":
    main()
