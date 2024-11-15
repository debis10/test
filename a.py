from docx import Document

def fill_template(data, template_path, output_path):
    # Load the template document
    doc = Document(template_path)

    # Replace placeholder text with actual data
    for paragraph in doc.paragraphs:
        for key, value in data.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, value)

    # Save the modified document
    doc.save(output_path)

# Data for each form
forms_data = [
    {
        "Full name": "John Doe",
        "USN": "1BY18CS101",
        "Address": "123 Main Street",
        "Mobile number": "9876543210",
        "Email": "john.doe@example.com",
        "Department": "Computer Science & Engineering",
        "Year of passing": "2021",
        "Field of interest": "MSc in Data Science",
        "Universities of interest": "University of Edinburgh, University of Oxford",
        "When do you wish to start your grad studies?": "2025",
        "TOEFL/IELTS Score": "7.5",
        "Current CGPA/Percentage": "8.5",
        "LoR’s by": "Dr. Smith - Professor",
        "Research/Work experience/Internship": "Internship at XYZ Corp - 6 months",
        "Budget": "₹ 20 Lac",
        "Signature of student": "John Doe",
    },
    # Add more dictionaries for other forms
]

# Generate the forms
template_path = 'HEF_Form.docx'
for i, data in enumerate(forms_data, start=1):
    output_path = f'HEF_Form_{i}.docx'
    fill_template(data, template_path, output_path)
    print(f"Generated form {i} at {output_path} okey")

