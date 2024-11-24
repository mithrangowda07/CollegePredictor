from django.shortcuts import render
from django.http import HttpResponse
import pandas as pd
import requests
from io import BytesIO
import csv

# Define the local file path for the Excel file
FILE_PATH = r"C:\Users\mithr\kcet_cutoff_viewer\cet_colg_data.xlsx"

def load_excel_file():
    """Load Excel file from local storage."""
    try:
        # Load the Excel file from the specified file path
        data = pd.read_excel(FILE_PATH, engine='openpyxl')
        return data, None
    except Exception as e:
        # Return None and the error message if an exception occurs
        return None, str(e)

def home(request):
    """Main page for KCET Cutoff Viewer."""
    df, error = load_excel_file()
    if error:
        return render(request, 'viewer/home.html', {'error': error})

    df.columns = df.columns.str.strip()
    categories = [col for col in df.columns if col not in ["College Code", "College Name", "Branch", "Branch code"]]
    colleges = df["College Name"].unique()
    branches = df["Branch"].unique()

    cutoff_rank = None
    if request.method == "POST":
        selected_category = request.POST.get("category")
        selected_college = request.POST.get("college")
        selected_branch = request.POST.get("branch")

        if selected_category and selected_college and selected_branch:
            filtered_data = df[
                (df["College Name"] == selected_college) &
                (df["Branch"] == selected_branch)
            ]
            if not filtered_data.empty:
                cutoff_rank = filtered_data.iloc[0][selected_category]
            else:
                error = "No data available for the selected options."
        else:
            error = "Please make valid selections for Category, College, and Branch."

    return render(request, 'viewer/home.html', {
        'categories': categories,
        'colleges': colleges,
        'branches': branches,
        'cutoff_rank': cutoff_rank,
    })

def download_csv(request):
    """Download selected colleges and cutoffs as CSV."""
    selected_colleges = request.session.get('selected_colleges', [])
    response = HttpResponse(content_type='text/csv')
    response['Content-Disposition'] = 'attachment; filename="selected_colleges.csv"'

    writer = csv.writer(response)
    writer.writerow(["College", "Branch", "Cutoff"])
    writer.writerows(selected_colleges)
    return response
