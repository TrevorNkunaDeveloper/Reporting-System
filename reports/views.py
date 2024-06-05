import pandas as pd
from django.shortcuts import render
from .forms import UploadFileForm
from django.http import HttpResponse, FileResponse
from django.core.files.storage import FileSystemStorage
from django.template.loader import get_template
from xhtml2pdf import pisa
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.units import inch

def load_transactions(file_path):
    transactions = pd.read_excel(file_path)
    print(transactions.columns)  # Print the column names to verify
    return transactions

def generate_report(transactions, start_date, end_date):
    start_date = pd.to_datetime(start_date)
    end_date = pd.to_datetime(end_date)
    
    transactions['dbCreateTime'] = pd.to_datetime(transactions['dbCreateTime'], errors='coerce')
    transactions['dbIssueTime'] = pd.to_datetime(transactions['dbIssueTime'], errors='coerce')
    
    filtered_transactions = transactions[
        (transactions['dbCreateTime'] >= start_date) & (transactions['dbCreateTime'] <= end_date)
    ].copy()

    total_revenue = filtered_transactions['dbPermitFee'].sum()
    number_of_permits_issued = filtered_transactions.shape[0]
    number_of_captured_permits = filtered_transactions[
        (filtered_transactions['dbIssueTime'] - filtered_transactions['dbCreateTime']).dt.total_seconds() <= 48 * 3600
    ].shape[0]
    percentage_captured = (number_of_captured_permits / number_of_permits_issued) * 100 if number_of_permits_issued > 0 else 0
    
    metrics = {
        "total_revenue": total_revenue,
        "number_of_permits_issued": number_of_permits_issued,
        "number_of_captured_permits": number_of_captured_permits,
        "percentage_captured": percentage_captured
    }
    
    filtered_transactions['dbCreateTime'] = filtered_transactions['dbCreateTime'].dt.strftime('%Y-%m-%d %H:%M:%S').fillna("")
    filtered_transactions['dbIssueTime'] = filtered_transactions['dbIssueTime'].dt.strftime('%Y-%m-%d %H:%M:%S').fillna("")

    report = filtered_transactions[[
        "dbPermitNo", "dbCreateTime", "dbAppDate", "dbIssueTime"
    ]].to_dict(orient='records')

    return report, metrics

def save_report_as_pdf(request, report, metrics):
    pdf_file = 'transactions_report.pdf'
    doc = SimpleDocTemplate(pdf_file, pagesize=letter)
    elements = []

    styles = getSampleStyleSheet()

    # Add the logo
    logo_path = 'static/images/logo.png'
    logo = Image(logo_path)
    logo.drawHeight = 1 * inch * logo.drawHeight / logo.drawWidth
    logo.drawWidth = 1 * inch
    logo.hAlign = 'CENTER'
    elements.append(logo)
    elements.append(Spacer(1, 0.25 * inch))

    header = Paragraph("Transaction Report", styles['Title'])
    elements.append(header)
    elements.append(Spacer(1, 0.5 * inch))

    metrics_paragraph = Paragraph(f"""
        <b>Metrics</b><br/>
        Total Revenue Collected (Rands): R {metrics['total_revenue']}<br/>
        Number of Permits Issued: {metrics['number_of_permits_issued']}<br/>
        Number of Permits Captured and Issued within 48 Hours: {metrics['number_of_captured_permits']}<br/>
        Percentage of Permits Captured and Issued within 48 Hours: {metrics['percentage_captured']:.2f}%
    """, styles['BodyText'])
    elements.append(metrics_paragraph)
    elements.append(Spacer(1, 0.5 * inch))

    data = [
        ['Permit Number', 'Date and Time Received', 'Date of Application', 'Date of Issue'],
    ]

    for item in report:
        data.append([item['dbPermitNo'], item['dbCreateTime'], item['dbAppDate'], item['dbIssueTime']])

    table = Table(data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))

    elements.append(table)
    doc.build(elements)

    return FileResponse(open(pdf_file, 'rb'), content_type='application/pdf')

def upload_file(request):
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            start_date = form.cleaned_data['start_date']
            end_date = form.cleaned_data['end_date']
            file = request.FILES['file']
            fs = FileSystemStorage()
            filename = fs.save(file.name, file)
            transactions = load_transactions(fs.path(filename))
            report, metrics = generate_report(transactions, start_date, end_date)
            
            if report is None:
                return render(request, 'reports/upload.html', {'form': form, 'error': metrics['error']})
            
            if 'save_pdf' in request.POST:
                return save_report_as_pdf(request, report, metrics)
            
            return render(request, 'reports/report.html', {'report': report, 'metrics': metrics})
    else:
        form = UploadFileForm()
    return render(request, 'reports/upload.html', {'form': form})
