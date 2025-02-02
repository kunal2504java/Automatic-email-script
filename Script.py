import smtplib
import pdfplumber
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from google.colab import files

# Function to send email
def send_email(to_email, subject, body, attachment_path):
    # Your email credentials
    from_email = "kunalpsingh25.com"  # Replace with your email
    password = "Kunal@programmer24"  # Replace with your password

    # Create a multipart message
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject

    # Attach the body text
    msg.attach(MIMEText(body, 'plain'))

    # Attach the resume PDF
    with open(attachment_path, 'rb') as attachment:
        part = MIMEApplication(attachment.read(), Name='Kunal-s-Resume-3.pdf')
        part['Content-Disposition'] = 'attachment; filename="Kunal-s-Resume-3.pdf"'
        msg.attach(part)

    # Send the email
    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(from_email, password)
            server.send_message(msg)
            print(f"Email sent to {to_email}")
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")

# Upload your resume file if not already uploaded (optional)
uploaded = files.upload()  # Uncomment if you need to upload again

# Path to your uploaded resume file in Colab
resume_path = 'Kunal-s-Resume-3.pdf'  # Ensure this matches the uploaded file name

# Read recruiter data from PDF file
pdf_file_path = 'hr_emails.pdf'  # Change this to your PDF file path

# Extract data from the PDF using pdfplumber
recruiter_data = []

with pdfplumber.open(pdf_file_path) as pdf:
    for page in pdf.pages:
        table = page.extract_table()
        if table:
            recruiter_data.extend(table[1:])  # Skip header row

# Iterate over each row in the extracted data and send emails
for row in recruiter_data:
    name = row[0]  # Assuming Name is in the first column
    hr_email = row[1]  # Assuming Email is in the second column
    title = row[2]  # Assuming Title is in the third column
    company_name = row[3]  # Assuming Company is in the fourth column
    
    subject = f"Application for Software Engineer Internship at {company_name}"
    
    # Personalize the body message using the name and title, and include company name
    body = f"""Dear {title} {name},

I hope this message finds you well. My name is Kunal Pratap Singh, and I am currently pursuing a Bachelor of Technology in Computer Science at G.L. Bajaj Institute of Technology, with an expected graduation date of August 2026. I am writing to express my interest in the Software Engineer Intern position at {company_name} as advertised.

I have a solid foundation in software development and engineering principles, bolstered by relevant coursework in Algorithms, Machine Learning, and Database Management. Additionally, my recent experiences include:

- **Co-Founder at Mon Zurich**: I launched a clothing brand and developed its website on Shopify, focusing on creating an engaging user experience with strong Front End design principles.

- **Cybersecurity Intern at Palo Alto Networks**: I completed a certification course in network security and participated in practical projects that honed my problem-solving skills.

- **Projects**: I have developed projects such as a Stock Price Prediction model using LSTM and a Transaction Management GUI in Java.

I am particularly drawn to {company_name} because of [specific reason related to the company or its projects/values]. I believe that my skills in Python, Java, and cloud technologies, combined with my proactive approach to learning and teamwork, make me a strong candidate for this role.

I have attached my resume for your review. I would be grateful for the opportunity to discuss how I can contribute to your team. Thank you for considering my application. I look forward to the possibility of speaking with you.

Best regards,

Kunal Pratap Singh  
Knowledge Park 3  
Greater Noida 201308  
Phone: 9005810309  
Email: kunalpsingh25@gmail.com  
LinkedIn: linkedin.com/in/kunal-pratap-singh  
GitHub: github.com/kunal2504java  
"""

    # Send the email (customize for each recipient)
    send_email(hr_email, subject, body.replace('[specific reason related to the company or its projects/values]', 'their innovative approach to technology'), resume_path)  # Replace with actual reason if known
