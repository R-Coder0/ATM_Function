import json
import os
import random
import smtplib
from win32com.client import Dispatch
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Path to the JSON file
JSON_FILE_PATH = os.path.join(os.path.dirname(__file__), 'account.json')

# Initialize text-to-speech
def speak(text):
    speaker = Dispatch("SAPI.SpVoice")
    speaker.Voice = speaker.GetVoices().Item(0)
    speaker.Speak(text)

# Load account data from JSON file
def load_account_data():
    if not os.path.exists(JSON_FILE_PATH):
        return {"accounts": []}
    with open(JSON_FILE_PATH, 'r') as file:
        return json.load(file)

# Save account data to JSON file
def save_account_data(data):
    with open(JSON_FILE_PATH, 'w') as file:
        json.dump(data, file, indent=4)

# Send OTP via email
def send_otp(email, otp):
    sender_email = "rcoder00@gmail.com"  # Replace with your email
    sender_password = "rnda bgep ybji vufr"  # Replace with your email password

    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = email
    message['Subject'] = 'Your OTP for Account Registration'

    body = f'Your OTP for registration is {otp}. Please do not share this with anyone.'
    message.attach(MIMEText(body, 'plain'))

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, email, message.as_string())
        server.quit()
        return True
    except Exception as e:
        print(f"Failed to send email. Error: {str(e)}")
        return False

# Register a new account
def register_account():
    speak("Please provide your account details for registration.")
    account_number = input("Enter new account number: ")
    name = input("Enter your name: ")
    account_type = input("Enter account type (Current/Savings): ")
    atm_pin = input("Enter new ATM pin: ")
    phone_number = input("Enter your phone number: ")
    email = input("Enter your email address: ")
    
    # Generate and send OTP
    otp = str(random.randint(100000, 999999))
    if send_otp(email, otp):
        speak("An OTP has been sent to your email address.")
        print("An OTP has been sent to your email address.")
    else:
        speak("Failed to send OTP. Registration cannot proceed.")
        print("Failed to send OTP. Registration cannot proceed.")
        return

    # Verify OTP
    entered_otp = input("Enter the OTP received: ")
    if entered_otp != otp:
        speak("Invalid OTP. Registration failed.")
        print("Invalid OTP. Registration failed.")
        return
    
    new_account = {
        "Account_number": account_number,
        "Name": name,
        "Account_type": account_type,
        "ATM_pin": atm_pin,
        "Account_balance": 0.0,  # Initial balance
        "Phone_number": phone_number,
        "Email": email
    }
    
    data = load_account_data()
    data['accounts'].append(new_account)
    save_account_data(data)
    
    speak("Registration successful.")
    print("Registration successful.")

# Handle transactions
def handle_transaction(account):
    atm_pin = input("Enter the ATM pin: ")
    speak("Enter your ATM pin.")
    
    if atm_pin != account["ATM_pin"]:
        speak("Invalid ATM pin. Transaction denied.")
        print("Invalid ATM pin. Transaction denied.")
        return
    
    account_type = input("Enter the account type (Current/Savings): ")
    speak("Enter the account type.")
    
    if account_type != account["Account_type"]:
        speak("Invalid account type. Transaction denied.")
        print("Invalid account type. Transaction denied.")
        return
    
    mobile_number = input("Enter your phone number: ")
    speak("Enter your phone number.")
    
    if mobile_number != account["Phone_number"]:
        speak("Invalid phone number. Transaction denied.")
        print("Invalid phone number. Transaction denied.")
        return
    
    try:
        withdraw_amount = float(input("Enter the amount to withdraw: "))
        speak("Enter the amount to withdraw.")
    except ValueError:
        speak("Invalid amount entered. Transaction denied.")
        print("Invalid amount entered. Transaction denied.")
        return
    
    if withdraw_amount > account["Account_balance"]:
        speak("Insufficient balance. Transaction denied.")
        print("Insufficient balance. Transaction denied.")
        return
    
    account["Account_balance"] -= withdraw_amount
    speak(f"Transaction successful. Your current balance is: {account['Account_balance']}.")
    print(f"Transaction successful. Your current balance is: {account['Account_balance']}")
    
    # Save updated account data
    account_data = load_account_data()
    for i, acc in enumerate(account_data['accounts']):
        if acc['Account_number'] == account['Account_number']:
            account_data['accounts'][i] = account
            break
    save_account_data(account_data)
    speak("Transaction successful.")
    print("Updated Account data is:")
    print(json.dumps(account, indent=4))

# Main function with menu options
def main():
    speak("Welcome to HDFC bank. Please choose an option.")
    print("Welcome to HDFC bank. Please choose an option:")
    print("1. Register a new account")
    print("2. Log in to an existing account")
    
    choice = input("Enter your choice (1/2): ")
    
    if choice == '1':
        register_account()
    elif choice == '2':
        account_data = load_account_data()
        account_number = input("Enter your account number: ")
        speak("Enter your account number.")
        
        # Find the account in the data
        account = next((acc for acc in account_data['accounts'] if acc['Account_number'] == account_number), None)
        
        if not account:
            speak("Account not found. Please register a new account.")
            print("Account not found. Please register a new account.")
        else:
            handle_transaction(account)
    else:
        speak("Invalid option selected.")
        print("Invalid option selected.")

if __name__ == "__main__":
    main()
