from datetime import datetime, timedelta

def emailDataGenerator(receiver, subject, sender):
    # Current date in YYYY-MM-DD format
    today = datetime.now()
    current_date = today.strftime("%Y-%m-%d")

    # Determine the current week's Monday and Sunday
    # weekday(): Monday=0, ..., Sunday=6
    monday = today - timedelta(days=today.weekday())
    sunday = monday + timedelta(days=6)

    # Format the week range: "Monday Dec 2, 2024 to Sunday Dec 8, 2024"
    current_week = f"{monday.strftime('%A %b %d, %Y')} to {sunday.strftime('%A %b %d, %Y')}"

    # Salutary remarks + contact info of sender
    salutations = f"Warm regards,\n{sender}\nYour Contact Info Here"

    return [receiver, subject, current_date, current_week, salutations]
