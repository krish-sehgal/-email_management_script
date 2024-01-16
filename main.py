import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def send_email(to_email, subject, body, sender_email, sender_password, sender_name):
    message = MIMEMultipart()
    message['From'] = f'{sender_name} <{sender_email}>'
    message['To'] = to_email
    message['Subject'] = subject
    message.attach(MIMEText(body, 'plain'))

    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, to_email, message.as_string())

def main():
    # Your email credentials
    sender_email = "your-email@gmail.com"
    sender_password = "aaaa bbbb cccc dddd"  # Replace with the app password you generated
    sender_name = "Your Name"  # Replace with your name

    # Read Excel file
    excel_file = 'data.xlsx'
    df = pd.read_excel(excel_file)

    choice = input("1 for SEND email\n2 for ADD email\nEnter your choice: ")
    choice = int(choice)

    if choice == 1:
        for index, row in df.iterrows():
            email = row['email']
            status = row['status']

            if status == 'no':
                subject = "Python Code"
                body = "Done bro!"
                send_email(email, subject, body, sender_email, sender_password, sender_name)
                df.loc[index, 'status'] = 'yes'
                print(f"Email sent to {email}")
    elif choice == 2:
        while True:
            new_email = input("Enter a new email (or press Enter to finish): ")
            if not new_email:
                break

            if new_email not in df['email'].values:
                df.loc[len(df.index)] = {'email': new_email, 'status': 'no'}
                print(f"Email '{new_email}' added successfully.")
            else:
                print("!! Email Already Exist !!")

    df.to_excel(excel_file, index=False, sheet_name='Sheet1')

if __name__ == "__main__":
    main()
