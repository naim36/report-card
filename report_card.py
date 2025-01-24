import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.pdfgen import canvas


try:
    df = pd.read_excel('student_scores.xlsx')
except Exception as e:
    print(f"Error reading the Excel file: {e}")
    exit()

# Check for missing or invalid data
if df.isnull().values.any():
    print("Warning: Missing data found in the Excel file.")
    df = df.dropna()  # Removing rows with missing data for this example

# Group data by student and calculate total and average scores
grouped = df.groupby(['Student ID', 'Name']).agg(
    total_score=('Subject Score', 'sum'),
    average_score=('Subject Score', 'mean')
).reset_index()

# Generate PDF report card for each student
for index, row in grouped.iterrows():
    student_id = row['Student ID']
    student_name = row['Name']
    total_score = row['total_score']
    average_score = row['average_score']

    # Create PDF canvas
    file_name = f'report_card_{student_id}.pdf'
    c = canvas.Canvas(file_name, pagesize=letter)
    c.setFont("Helvetica", 12)

    # Title and student info
    c.drawString(100, 750, f"Report Card for {student_name} (ID: {student_id})")
    c.drawString(100, 730, f"Total Score: {total_score}")
    c.drawString(100, 710, f"Average Score: {average_score:.2f}")

    # Create table for subject-wise scores
    c.drawString(100, 670, "Subject-wise Scores:")
    
    # Prepare data for table
    student_scores = df[df['Student ID'] == student_id]
    y_position = 650
    for i, score_row in student_scores.iterrows():
        subject = score_row['Subject']
        score = score_row['Subject Score']
        c.drawString(100, y_position, f"{subject}: {score}")
        y_position -= 20
    
    # Save PDF
    c.save()

print("Report cards generated successfully.")
