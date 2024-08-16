import json
import os
import random
import smtplib
from win32com.client import Dispatch
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import tkinter as tk
from tkinter import messagebox, simpledialog

JSON_FILE_PATH = os.path.join(os.path.dirname(__file__), 'account.json')

def speak(text):
    speaker = Dispatch("SAPI.SpVoice")
    speaker.Voice = speaker.GetVoices().Item(0)
    speaker.Speak(text)

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
    sender_email = "rcoder00@gmail.com" 
    sender_password = "rnda bgep ybji vufr" 

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
        messagebox.showerror("Error", f"Failed to send email. Error: {str(e)}")
        return False

# Register a new account
def register_account():
    account_number = simpledialog.askstring("Account Number", "Enter new account number:")
    name = simpledialog.askstring("Name", "Enter your name:")
    account_type = simpledialog.askstring("Account Type", "Enter account type (Current/Savings):")
    atm_pin = simpledialog.askstring("ATM PIN", "Enter new ATM pin:")
    phone_number = simpledialog.askstring("Phone Number", "Enter your phone number:")
    email = simpledialog.askstring("Email", "Enter your email address:")

    otp = str(random.randint(100000, 999999))
    if send_otp(email, otp):
        speak("An OTP has been sent to your email address.")
        messagebox.showinfo("OTP Sent", "An OTP has been sent to your email address.")
    else:
        speak("Failed to send OTP. Registration cannot proceed.")
        return

    entered_otp = simpledialog.askstring("OTP Verification", "Enter the OTP received:")
    if entered_otp != otp:
        speak("Invalid OTP. Registration failed.")
        messagebox.showerror("Error", "Invalid OTP. Registration failed.")
        return

    initial_deposit = float(simpledialog.askstring("Initial Deposit", "Enter initial deposit amount:"))
    speak("Enter initial deposit amount.")

    new_account = {
        "Account_number": account_number,
        "Name": name,
        "Account_type": account_type,
        "ATM_pin": atm_pin,
        "Account_balance": initial_deposit,
        "Phone_number": phone_number,
        "Email": email
    }
    
    data = load_account_data()
    data['accounts'].append(new_account)
    save_account_data(data)
    
    speak("Registration successful.")
    messagebox.showinfo("Success", "Registration successful.")

# Handle transactions
def handle_transaction(account):
    atm_pin = simpledialog.askstring("ATM PIN", "Enter the ATM pin:")
    if atm_pin != account["ATM_pin"]:
        speak("Invalid ATM pin. Transaction denied.")
        messagebox.showerror("Error", "Invalid ATM pin. Transaction denied.")
        return
    
    account_type = simpledialog.askstring("Account Type", "Enter the account type (Current/Savings):")
    if account_type != account["Account_type"]:
        speak("Invalid account type. Transaction denied.")
        messagebox.showerror("Error", "Invalid account type. Transaction denied.")
        return
    
    mobile_number = simpledialog.askstring("Phone Number", "Enter your phone number:")
    if mobile_number != account["Phone_number"]:
        speak("Invalid phone number. Transaction denied.")
        messagebox.showerror("Error", "Invalid phone number. Transaction denied.")
        return
    
    choice = messagebox.askquestion("Transaction", "Would you like to withdraw money?")
    
    if choice == 'yes':
        withdraw_amount = float(simpledialog.askstring("Withdraw Amount", "Enter the amount to withdraw:"))
        if withdraw_amount > account["Account_balance"]:
            speak("Insufficient balance. Transaction denied.")
            messagebox.showerror("Error", "Insufficient balance. Transaction denied.")
            return
        account["Account_balance"] -= withdraw_amount

    else:
        deposit_amount = float(simpledialog.askstring("Deposit Amount", "Enter the amount to deposit:"))
        account["Account_balance"] += deposit_amount
    
    account_data = load_account_data()
    for i, acc in enumerate(account_data['accounts']):
        if acc['Account_number'] == account['Account_number']:
            account_data['accounts'][i] = account
            break
    save_account_data(account_data)
    
    speak("Transaction successful.")
    messagebox.showinfo("Success", f"Transaction successful. Your current balance is: {account['Account_balance']}.")

# Main function with UI
def main():
    root = tk.Tk()
    root.title("HDFC Bank Application")

    tk.Label(root, text="Welcome to HDFC Bank", font=("Arial", 16)).pack(pady=10)
    tk.Label(root, text="Please choose an option", font=("Arial", 14)).pack(pady=10)

    tk.Button(root, text="Register a New Account", command=register_account, font=("Arial", 12)).pack(pady=10)
    tk.Button(root, text="Log in to Existing Account", command=lambda: login(root), font=("Arial", 12)).pack(pady=10)

    root.mainloop()

def login(root):
    account_data = load_account_data()
    account_number = simpledialog.askstring("Login", "Enter your account number:")

    account = next((acc for acc in account_data['accounts'] if acc['Account_number'] == account_number), None)
    
    if not account:
        speak("Account not found. Please register a new account.")
        messagebox.showerror("Error", "Account not found. Please register a new account.")
    else:
        handle_transaction(account)

if __name__ == "__main__":
    main()
