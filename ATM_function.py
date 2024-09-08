import json
import os
import random
import smtplib
from win32com.client import Dispatch
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import tkinter as tk
from tkinter import simpledialog, messagebox

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

def save_account_data(data):
    with open(JSON_FILE_PATH, 'w') as file:
        json.dump(data, file, indent=4)

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
        print(f"Error sending email: {e}")
        return False

def generate_account_number():
    return "ACC" + str(random.randint(10000000, 99999999))

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
    except Exception as e:
        print(f"Error sending email: {e}")

def register_account(frame, root):
    for widget in frame.winfo_children():
        widget.destroy()

    tk.Label(frame, text="Register New Account", font=("Arial", 20, "bold")).pack(pady=10)
    
    tk.Label(frame, text="Account Number").pack(pady=5)
    account_number = generate_account_number()
    tk.Label(frame, text=account_number, font=("Arial", 12, "bold")).pack()

    tk.Label(frame, text="Enter Name").pack(pady=5)
    name_entry = tk.Entry(frame, width=40)
    name_entry.pack()

    tk.Label(frame, text="Select Account Type").pack(pady=5)
    account_type_var = tk.StringVar(value="Savings")
    
    account_type_frame = tk.Frame(frame)
    account_type_frame.pack(pady=5)
    
    tk.Radiobutton(account_type_frame, text="Savings", variable=account_type_var, value="Savings").pack(side="left", padx=10)
    tk.Radiobutton(account_type_frame, text="Current", variable=account_type_var, value="Current").pack(side="left", padx=10)

    tk.Label(frame, text="Enter ATM PIN").pack(pady=5)
    atm_pin_entry = tk.Entry(frame, show="*", width=40)
    atm_pin_entry.pack()

    tk.Label(frame, text="Enter Phone Number").pack(pady=5)
    phone_number_entry = tk.Entry(frame, width=40)
    phone_number_entry.pack()

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
                    # Create a custom window for entering the initial deposit
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
                            tk.Label(deposit_window, text="Please enter a valid amount.", fg="red", font=("Arial", 12)).pack()

                    tk.Button(deposit_window, text="Confirm", command=confirm_deposit, width=15, bg="lightgreen", font=("Arial", 12)).pack(pady=20)

                else:
                    print("OTP did not match")  # Debugging statement
                    speak("Invalid OTP. Registration failed.")

            verify_button = tk.Button(frame, text="Verify OTP", command=verify_otp, width=20, bg="lightblue")
            verify_button.pack(pady=10)

        else:
            speak("Failed to send OTP.")

    submit_button = tk.Button(frame, text="Submit", command=submit_registration, width=20, bg="lightgreen")
    submit_button.pack(pady=20)

def handle_transaction(account, frame, root):
    for widget in frame.winfo_children():
        widget.destroy()

    balance_label = tk.Label(frame, text=f"Account Balance: ₹{account['Account_balance']:.2f}", font=("Arial", 14))
    balance_label.pack(pady=10)

    def update_balance_label():
        balance_label.config(text=f"Account Balance: ₹{account['Account_balance']:.2f}")

    def withdraw_money():
        # Create a new window for the withdrawal input
        withdraw_window = tk.Toplevel(root)
        withdraw_window.title("Withdraw Money")
        withdraw_window.geometry("400x200")
        withdraw_window.configure(bg='#f0f8ff')

        tk.Label(withdraw_window, text="Enter the amount to withdraw:", font=("Arial", 14)).pack(pady=20)

        withdraw_entry = tk.Entry(withdraw_window, font=("Arial", 14), width=20)
        withdraw_entry.pack(pady=10)

        def confirm_withdraw():
            try:
                withdraw_amount = float(withdraw_entry.get())
                if withdraw_amount > account["Account_balance"]:
                    speak("Insufficient balance.")
                    messagebox.showerror("Error", "Insufficient balance. Please enter a valid amount.")
                else:
                    account["Account_balance"] -= withdraw_amount
                    save_account_data(load_account_data())  # Save the updated balance
                    update_balance_label()  # Refresh the balance display
                    speak("Transaction successful.")
                    messagebox.showinfo("Success", "Withdrawal completed successfully.")
                    withdraw_window.destroy()
            except ValueError:
                messagebox.showerror("Error", "Please enter a valid amount.")

        tk.Button(withdraw_window, text="Confirm Withdraw", command=confirm_withdraw, width=15, bg="lightgreen", font=("Arial", 12)).pack(pady=20)

    def deposit_money():
        # Create a new window for the deposit input
        deposit_window = tk.Toplevel(root)
        deposit_window.title("Deposit Money")
        deposit_window.geometry("400x200")
        deposit_window.configure(bg='#f0f8ff')

        tk.Label(deposit_window, text="Enter the amount to deposit:", font=("Arial", 14)).pack(pady=20)

        deposit_entry = tk.Entry(deposit_window, font=("Arial", 14), width=20)
        deposit_entry.pack(pady=10)

        def confirm_deposit():
            try:
                deposit_amount = float(deposit_entry.get())
                account["Account_balance"] += deposit_amount
                save_account_data(load_account_data())  # Save the updated balance
                update_balance_label()  # Refresh the balance display
                speak("Transaction successful.")
                messagebox.showinfo("Success", "Deposit completed successfully.")
                deposit_window.destroy()
            except ValueError:
                messagebox.showerror("Error", "Please enter a valid amount.")

        tk.Button(deposit_window, text="Confirm Deposit", command=confirm_deposit, width=15, bg="lightgreen", font=("Arial", 12)).pack(pady=20)

    tk.Button(frame, text="Withdraw Money", command=withdraw_money, width=20, bg="#f4a460", font=("Arial", 14)).pack(pady=10)
    tk.Button(frame, text="Deposit Money", command=deposit_money, width=20, bg="#90ee90", font=("Arial", 14)).pack(pady=10)

def login(frame, root):
    for widget in frame.winfo_children():
        widget.destroy()

    tk.Label(frame, text="Login to Account", font=("Arial", 20, "bold")).pack(pady=10)
    
    tk.Label(frame, text="Enter Account Number").pack(pady=5)
    account_number_entry = tk.Entry(frame, width=40)
    account_number_entry.pack()

    # Function to capitalize alphabetic characters as the user types
    def capitalize_account_number(event):
        current_text = account_number_entry.get()
        account_number_entry.delete(0, tk.END)
        account_number_entry.insert(0, current_text.upper())

    # Bind the key release event to the capitalize_account_number function
    account_number_entry.bind("<KeyRelease>", capitalize_account_number)

    tk.Label(frame, text="Enter ATM PIN").pack(pady=5)
    atm_pin_entry = tk.Entry(frame, show="*", width=40)
    atm_pin_entry.pack()

    def verify_login():
        account_number = account_number_entry.get()
        atm_pin = atm_pin_entry.get()
        data = load_account_data()
        account = next((acc for acc in data['accounts'] if acc['Account_number'] == account_number and acc['ATM_pin'] == atm_pin), None)
        if account:
            handle_transaction(account, frame, root)
        else:
            speak("Invalid credentials. Please try again.")

    tk.Button(frame, text="Login", command=verify_login, width=20, bg="lightblue", font=("Arial", 14)).pack(pady=10)


def show_main_menu(frame, root):
    for widget in frame.winfo_children():
        widget.destroy()

    tk.Label(frame, text="Welcome to HDFC Bank", font=("Arial", 24, "bold")).pack(pady=20)
    tk.Label(frame, text="Please choose an option", font=("Arial", 16)).pack(pady=10)

    tk.Button(frame, text="Register a New Account", command=lambda: register_account(frame, root), width=25, bg='#90ee90', fg='#333', font=("Arial", 14)).pack(pady=10)
    tk.Button(frame, text="Log in to Existing Account", command=lambda: login(frame, root), width=25, bg='#add8e6', fg='#333', font=("Arial", 14)).pack(pady=10)

def main():
    global root
    root = tk.Tk()
    root.title("HDFC Bank Application")
    root.geometry("700x600")
    root.configure(bg='#f0f8ff')

    frame = tk.Frame(root, bg='#f0f8ff')
    frame.pack(expand=True, fill=tk.BOTH)

    show_main_menu(frame, root)
    
    root.mainloop()

if __name__ == "__main__":
    main()
