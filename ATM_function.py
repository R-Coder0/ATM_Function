import json
from win32com.client import Dispatch
import os

# Path to the JSON file
JSON_FILE_PATH = os.path.join(os.path.dirname(__file__), 'account.json')

# Initialize text-to-speech
def speak(text):
    speaker = Dispatch("SAPI.SpVoice")
    speaker.Voice = speaker.GetVoices().Item(0)
    speaker.Speak(text)

# Load account data from JSON file
def load_account_data():
    with open(JSON_FILE_PATH, 'r') as file:
        return json.load(file)

# Save account data to JSON file
def save_account_data(data):
    with open(JSON_FILE_PATH, 'w') as file:
        json.dump(data, file, indent=4)

# Register a new account
def register_account():
    speak("Please provide your account details for registration.")
    account_number = input("Enter new account number: ")
    name = input("Enter your name: ")
    account_type = input("Enter account type (Current/Savings): ")
    atm_pin = input("Enter new ATM pin: ")
    phone_number = input("Enter your phone number: ")
    
    new_account = {
        "Account_number": account_number,
        "Name": name,
        "Account_type": account_type,
        "ATM_pin": atm_pin,
        "Account_balance": 0.0,  # Initial balance
        "Phone_number": phone_number
    }
    
    data = load_account_data()
    data['accounts'].append(new_account)
    save_account_data(data)
    
    speak("Registration successful.")
    print("Registration successful.")

# Main function to handle transactions
def main():
    speak("Welcome to HDFC bank. How may I help you today?")
    account_data = load_account_data()
    
    account_number = input("Enter your account number: ")
    speak("Enter your account number.")
    
    # Find the account in the data
    account = next((acc for acc in account_data['accounts'] if acc['Account_number'] == account_number), None)
    
    if not account:
        speak("Account not found. Would you like to register? (yes/no)")
        if input("Account not found. Would you like to register? (yes/no): ").lower() == 'yes':
            register_account()
        else:
            speak("Transaction denied.")
            print("Transaction denied.")
        return
    
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
    save_account_data(account_data)
    speak("Transaction successful.")
    print("Updated Account data is:")
    print(json.dumps(account, indent=4))

if __name__ == "__main__":
    main()
