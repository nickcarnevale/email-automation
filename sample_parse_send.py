import pandas as pd
import smtplib
from email.message import EmailMessage

def parse_excel(filename, category):
    # Load the Excel file into a pandas DataFrame
    try:
        # Load the Excel file into a pandas DataFrame
        df = pd.read_excel(filename)
    except ImportError:
        raise ImportError("Error opening excel file")

    #for case insensitivity
    df['Category'] = df['Category'].str.lower()

    #check if category is in Excel File
    if category.lower() not in df['Category'].unique():
        print("Category not found in the Excel file.")
        return []

    #filter the data based on if the category matches the given category
    filtered_df = df[df['Category'].str.lower() == category.lower()]

    # Extract the email and name columns
    email_column = filtered_df['Email']
    name_column = filtered_df['Name']

    # Create an empty list to store the email/name pairs
    email_name_pairs = []

    # Iterate over the rows and retrieve the email/name pairs
    for i in range(len(filtered_df)):
        email = email_column.iloc[i]
        name = name_column.iloc[i]
        email_name_pairs.append((email, name))

    # Return the email/name pairs
    return email_name_pairs


# Send email to each email address
def send_email(email, name):
    # Set up the email message
    msg = EmailMessage()
    msg['Subject'] = f'Testing {name}!'
    msg['From'] = 'devolatesterl@gmail.com'
    msg['To'] = email
    msg.set_content(f"Hi {name}! This is a message from Devola Tech.")

    # Set up the SMTP server and send the email
    with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.login('devolatester@gmail.com', 'xarkrrlckjffifmq')
        smtp.send_message(msg)


# Test the parser with user input for category
filename = 'sampledatabase.xlsx'
category = input("Enter the category: ")
pairs = parse_excel(filename, category)

for email, name in pairs:
    send_email(email, name)
    print(f"Email sent to {email} with 'Hi {name}!'")
