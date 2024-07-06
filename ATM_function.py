import json

account_data = {
    "Account_number" : '123456789456',
    "Name":'Rishabh Pratap',
    "Account_type":'Savings',
    "ATM_pin" : '5486',
    "Account_balance": 4000.0,
    "Phone_number" : '987654321'
}

from win32com.client import Dispatch

speak = Dispatch("SAPI.SpVoice")
speak.Voice = speak.GetVoices().Item(0)
speak.Speak("Wecome to HDFC bank how may i help you today")

account_number = input("Enter you account number : ")
from win32com.client import Dispatch

speak = Dispatch("SAPI.SpVoice")
speak.Voice = speak.GetVoices().Item(0)
speak.Speak("Enter you account number")

if account_number != account_data['Account_number']:
    print("Invalid account number. Transaction denied!!")
else:
    atm_pin = input("Enter the ATM pin : ")
    from win32com.client import Dispatch

    speak = Dispatch("SAPI.SpVoice")
    speak.Voice = speak.GetVoices().Item(0)
    speak.Speak("Enter you atm pin")

    if atm_pin != account_data["ATM_pin"] :
       print("Invalid ATM pin. Transaction denied!!")
       speak = Dispatch("SAPI.SpVoice")
       speak.Voice = speak.GetVoices().Item(0)
       speak.Speak("Transaction Denied")
    else:
        account_type = input("Enter the account type Current/Savings : ")
        if account_type != account_data["Account_type"]:
            print("Invalid account type. Transaction denied!!")
        else:
            mobile_number = input("Enter your phone number : ")
            if mobile_number != account_data["Phone_number"]:
                print("Invalid phone number. Transaction denied!!")
            else:
                try:
                    withdraw_amount = float(input("Enter your amount to withdraw : "))
                except ValueError:
                    print("Invalid amount enter. Transaction denied")
                else:
                    if withdraw_amount > account_data["Account_balance"]:
                        print("Insufficient balance. Transaction denied!!")
                        from win32com.client import Dispatch

                        speak = Dispatch("SAPI.SpVoice")
                        speak.Voice = speak.GetVoices().Item(0)
                        speak.Speak("Insuficent Balance")
                    else:
                        account_data["Account_balance"] -= withdraw_amount
                        print(f"Transaction successfull. Your current balance is : {account_data['Account_balance']}")
                        from win32com.client import Dispatch

                        speak = Dispatch("SAPI.SpVoice")
                        speak.Voice = speak.GetVoices().Item(0)
                        speak.Speak("Transaction Successfull.")
                        
                         
                        
                        update_account_data = json.dumps(account_data, indent=4)
                        print(f"Updated Account data is :\n {update_account_data}")
                        
                    
        