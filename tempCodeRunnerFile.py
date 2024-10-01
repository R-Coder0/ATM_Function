import json
import os
import random
import smtplib
from win32com.client import Dispatch
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import tkinter as tk
from tkinter import simpledialog, messagebox

# Define the path to your JSON file
JSON_FILE_PATH = os.path.join(os.path.dirname(__file__), 'account.json')

# Function to make the program speak using AI voice
def speak(text):
    speaker = Dispatch("SAPI.SpVoice")
    speaker.Voice = speaker.GetVoices().Item(0)
    speaker.Speak(text)

# Load account data from a JSON file
def load_account_data():
    if not os.path.exists(JSON_FILE_PATH):
        return {"accounts": []}
    with open(JSON_FILE_PATH, 'r') as file:
        return json.load(file)

# Save account data to a JSON file
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
        speak("OTP has been sent successfully to the email.")
        return True
    except Exception as e:
        print(f"Error sending email: {e}")
        speak("Failed to send OTP. Please try again.")
        return False

# Generate a random account number
def generate_account_number():
    return "ACC" + str(random.randint(10000000, 99999999))

# Send registration details via email
def send_registration_details(email, account_details):
    sender_email = "rcoder00@gmail.com"
    sender_password = "rnda bgep ybji vufr"

    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = email
    message['Subject'] = 'Your HDFC Bank Account Registration Details'
    
    body = (
        f"Dear {account_details['Name']},\n\n"
        f"Thank you for registering with HDFC Bank. Here are your account details:\n"
        f"Account Number: {account_details['Account_number']}\n"
        f"Account Type: {account_details['Account_type']}\n"
        f"ATM PIN: {account_details['ATM_pin']}\n"
        f"Account Balance: ₹{account_details['Account_balance']:.2f}\n"
        f"Phone Number: {account_details['Phone_number']}\n"
        f"Email: {account_details['Email']}\n\n"
        f"Please keep this information safe and do not share your ATM PIN with anyone.\n\n"
        f"Best regards,\n"
        f"HDFC Bank"
    )
    message.attach(MIMEText(body, 'plain'))

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, email, message.as_string())
        server.quit()
        speak("Registration details have been sent to your email.")
    except Exception as e:
        print(f"Error sending email: {e}")
        speak("Failed to send registration details.")

# Register new account
def register_account(frame, root):
    for widget in frame.winfo_children():
        widget.destroy()

    speak("Registering a new account.")
    
    tk.Label(frame, text="Register New Account", font=("Arial", 20, "bold")).pack(pady=10)
    
    tk.Label(frame, text="Account Number").pack(pady=5)
    account_number = generate_account_number()
    tk.Label(frame, text=account_number, font=("Arial", 12, "bold")).pack()

    speak("Please enter your name.")
    tk.Label(frame, text="Enter Name").pack(pady=5)
    name_entry = tk.Entry(frame, width=40)
    name_entry.pack()

    speak("Please select the account type.")
    tk.Label(frame, text="Select Account Type").pack(pady=5)
    account_type_var = tk.StringVar(value="Savings")
    
    account_type_frame = tk.Frame(frame)
    account_type_frame.pack(pady=5)
    
    tk.Radiobutton(account_type_frame, text="Savings", variable=account_type_var, value="Savings").pack(side="left", padx=10)
    tk.Radiobutton(account_type_frame, text="Current", variable=account_type_var, value="Current").pack(side="left", padx=10)

    speak("Please enter a 4-digit ATM PIN.")
    tk.Label(frame, text="Enter ATM PIN").pack(pady=5)
    atm_pin_entry = tk.Entry(frame, show="*", width=40)
    atm_pin_entry.pack()

    speak("Please enter your phone number.")
    tk.Label(frame, text="Enter Phone Number").pack(pady=5)
    phone_number_entry = tk.Entry(frame, width=40)
    phone_number_entry.pack()

    speak("Please enter your email address.")
    tk.Label(frame, text="Enter Email Address").pack(pady=5)
    email_entry = tk.Entry(frame, width=40)
    email_entry.pack()

    def submit_registration():
        name = name_entry.get().title()
        account_type = account_type_var.get()
        atm_pin = atm_pin_entry.get()
        phone_number = phone_number_entry.get()
        email = email_entry.get()

        otp = str(random.randint(100000, 999999))
        speak("An OTP is being generated.")
        print(f"Generated OTP: {otp}")  # Debugging statement
        if send_otp(email, otp):
            speak("An OTP has been sent to your email address.")
            otp_label = tk.Label(frame, text="Enter OTP sent to your email")
            otp_label.pack(pady=5)
            otp_entry = tk.Entry(frame, width=40)
            otp_entry.pack()

            def verify_otp():
                entered_otp = otp_entry.get()
                print(f"Entered OTP: {entered_otp}")  # Debugging statement
                if entered_otp == otp:
                    print("OTP matched")  # Debugging statement
                    speak("OTP matched successfully.")
                    
                    deposit_window = tk.Toplevel(root)
                    deposit_window.title("Initial Deposit")
                    deposit_window.geometry("400x250")
                    deposit_window.configure(bg='#f0f8ff')

                    tk.Label(deposit_window, text="Enter Initial Deposit Amount:", font=("Arial", 14)).pack(pady=20)

                    deposit_entry = tk.Entry(deposit_window, font=("Arial", 14), width=20)
                    deposit_entry.pack(pady=10)

                    def confirm_deposit():
                        try:
                            initial_deposit = float(deposit_entry.get())
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
                            send_registration_details(email, new_account)
                            speak("Registration successful.")
                            messagebox.showinfo("Success", "Registration completed successfully.")
                            deposit_window.destroy()
                            show_main_menu(frame, root)
                        except ValueError:
                            speak("Invalid amount entered. Please try again.")
                            tk.Label(deposit_window, text="Please enter a valid amount.", fg="red", font=("Arial", 12)).pack()

                    tk.Button(deposit_window, text="Confirm", command=confirm_deposit, width=15, bg="lightgreen", font=("Arial", 12)).pack(pady=20)

                else:
                    speak("Invalid OTP. Please try again.")
                    print("OTP did not match")  # Debugging statement

            verify_button = tk.Button(frame, text="Verify OTP", command=verify_otp, width=20, bg="lightblue")
            verify_button.pack(pady=10)

        else:
            speak("Failed to send OTP.")

    submit_button = tk.Button(frame, text="Submit", command=submit_registration, width=20, bg="lightgreen")
    submit_button.pack(pady=20)

# Handle transaction (Withdraw and Deposit)
def handle_transaction(account, frame, root):
    for widget in frame.winfo_children():
        widget.destroy()

    speak(f"Your current balance is {account['Account_balance']:.2f} rupees.")
    
    balance_label = tk.Label(frame, text=f"Account Balance: ₹{account['Account_balance']:.2f}", font=("Arial", 14))
    balance_label.pack(pady=10)

    def update_balance_label():
        balance_label.config(text=f"Account Balance: ₹{account['Account_balance']:.2f}")

    def withdraw_money():
        withdraw_window = tk.Toplevel(root)
        withdraw_window.title("Withdraw Money")
        withdraw_window.geometry("400x200")
        withdraw_window.configure(bg='#f0f8ff')

        tk.Label(withdraw_window, text="Enter Withdrawal Amount:", font=("Arial", 14)).pack(pady=10)
        withdraw_entry = tk.Entry(withdraw_window, font=("Arial", 14), width=20)
        withdraw_entry.pack(pady=10)

        def confirm_withdraw():
            try:
                amount = float(withdraw_entry.get())
                if amount <= 0:
                    speak("Invalid amount. Please enter a positive number.")
                    raise ValueError("Invalid amount")
                if amount > account['Account_balance']:
                    speak("Insufficient balance.")
                    messagebox.showwarning("Error", "Insufficient balance")
                else:
                    account['Account_balance'] -= amount
                    speak(f"Withdrawal of {amount:.2f} rupees is successful.")
                    data = load_account_data()
                    for acc in data['accounts']:
                        if acc['Account_number'] == account['Account_number']:
                            acc['Account_balance'] = account['Account_balance']
                            break
                    save_account_data(data)
                    withdraw_window.destroy()
                    update_balance_label()
                    speak(f"Your updated balance is {account['Account_balance']:.2f} rupees.")
            except ValueError:
                speak("Invalid input. Please try again.")
                messagebox.showwarning("Error", "Invalid input")

        tk.Button(withdraw_window, text="Withdraw", command=confirm_withdraw, width=20, bg="lightgreen").pack(pady=10)

    def deposit_money():
        deposit_window = tk.Toplevel(root)
        deposit_window.title("Deposit Money")
        deposit_window.geometry("400x200")
        deposit_window.configure(bg='#f0f8ff')

        tk.Label(deposit_window, text="Enter Deposit Amount:", font=("Arial", 14)).pack(pady=10)
        deposit_entry = tk.Entry(deposit_window, font=("Arial", 14), width=20)
        deposit_entry.pack(pady=10)

        def confirm_deposit():
            try:
                amount = float(deposit_entry.get())
                if amount <= 0:
                    speak("Invalid amount. Please enter a positive number.")
                    raise ValueError("Invalid amount")
                account['Account_balance'] += amount
                speak(f"Deposit of {amount:.2f} rupees is successful.")
                data = load_account_data()
                for acc in data['accounts']:
                    if acc['Account_number'] == account['Account_number']:
                        acc['Account_balance'] = account['Account_balance']
                        break
                save_account_data(data)
                deposit_window.destroy()
                update_balance_label()
                speak(f"Your updated balance is {account['Account_balance']:.2f} rupees.")
            except ValueError:
                speak("Invalid input. Please try again.")
                messagebox.showwarning("Error", "Invalid input")

        tk.Button(deposit_window, text="Deposit", command=confirm_deposit, width=20, bg="lightgreen").pack(pady=10)

    tk.Button(frame, text="Withdraw", command=withdraw_money, width=20, bg="lightblue").pack(pady=10)
    tk.Button(frame, text="Deposit", command=deposit_money, width=20, bg="lightgreen").pack(pady=10)
    tk.Button(frame, text="Back to Main Menu", command=lambda: show_main_menu(frame, root), width=20, bg="lightgray").pack(pady=10)

# Show main menu
def show_main_menu(frame, root):
    for widget in frame.winfo_children():
        widget.destroy()

    speak("Welcome to HDFC Bank.")
    
    tk.Label(frame, text="HDFC Bank", font=("Arial", 24, "bold")).pack(pady=10)
    
    def login():
        account_number = simpledialog.askstring("Login", "Enter your account number:")
        data = load_account_data()
        account = next((acc for acc in data['accounts'] if acc['Account_number'] == account_number), None)
        if account:
            speak("Login successful.")
            handle_transaction(account, frame, root)
        else:
            speak("Account not found. Please try again.")
            messagebox.showerror("Error", "Account not found")

    tk.Button(frame, text="Register New Account", command=lambda: register_account(frame, root), width=30, height=2, bg="lightblue").pack(pady=10)
    tk.Button(frame, text="Login", command=login, width=30, height=2, bg="lightgreen").pack(pady=10)
    tk.Button(frame, text="Exit", command=root.quit, width=30, height=2, bg="lightgray").pack(pady=10)

# Initialize GUI
root = tk.Tk()
root.title("HDFC Bank")
root.geometry("600x400")

frame = tk.Frame(root)
frame.pack(expand=True, fill="both")

show_main_menu(frame, root)
root.mainloop()
