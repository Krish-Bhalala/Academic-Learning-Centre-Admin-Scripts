import smtplib
from email.message import EmailMessage

def sendEmail(emailData, receiverEmail):
    # emailData format: [receiver email, subject, body, email closing, sender email]

    data_receiver, subject, body, closing, sender = emailData
    final_receiver = receiverEmail if receiverEmail else data_receiver

    msg = EmailMessage()
    msg['From'] = sender
    msg['To'] = final_receiver
    msg['Subject'] = subject
    full_message = f"{body}\n\n{closing}"
    msg.set_content(full_message)

    smtp_host = 'smtp.gmail.com'
    smtp_port = 587
    sender_email = sender
    sender_password = "<YOUR_EMAIL_PASSWORD>"  # Replace with a secure method

    try:
        with smtplib.SMTP(smtp_host, smtp_port) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
            return True
    except Exception as e:
        print(f"Error sending email: {e}")
        return "email not sent"
