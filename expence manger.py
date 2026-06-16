import calendar
import datetime
import speech_recognition as re
import win32com.client  

expenselist = ["food", "transportation", "entertainment", "utilities", "healthcare", "education", "shopping", "travel", "miscellaneous", "rent", "groceries", "dining out", "subscriptions", "insurance", "gifts", "charity", "personal care", "fitness", "hobbies"]
speaker = win32com.client.Dispatch("SAPI.SpVoice")
recognizer = re.Recognizer() 

totalexpence = 0.0
expense_history = []
expense_category_history = []
expence_history = []
expence_date_history = []

def speak_and_listen(text_to_speak):
    print(text_to_speak)
    speaker.Speak(text_to_speak) 
    while True:
        try:
            with re.Microphone() as source:
                print("listening...")
                audio = recognizer.listen(source, timeout=4, phrase_time_limit=4)
                text = recognizer.recognize_google(audio)
                print("You said:", text)
                return text  
        except Exception:
            print("Could not hear clearly. Retrying...")

def save_expense_to_csv(date, expense, amount):
    open("expenses.csv", "a").write(f"{date},{expense},{amount}\n")

def add_expense(choice):
    global totalexpence
    if choice == 1:
        date = speak_and_listen("Enter the date in YYYY-MM-DD and HH:MM format:")
        if date.lower() == "stop": return "stop"
        expense = speak_and_listen("Enter the type of expense:")
        if expense.lower() == "stop": return "stop"
        if expense not in expenselist: expenselist.append(expense)
        amount = speak_and_listen("Enter the amount of expense:")
        if amount.lower() == "stop": return "stop"
    else:
        date = input("Enter the date in YYYY-MM-DD and HH:MM format (or type 'stop'): ")
        if date.lower() == "stop": return "stop"
        expense = input("enter the type of expense: ")
        if expense.lower() == "stop": return "stop"
        amount = input("Enter the amount of expense: ")
        if amount.lower() == "stop": return "stop"

    try:
        amount_numeric = float(amount)
    except ValueError:
        print("Error: Amount must be a number! Restarting row...")
        if choice == 1: speaker.Speak("Invalid number. Re-entering row.")
        return "continue"

    if choice == 1: speaker.Speak("Expense recorded successfully!")
    dailyexpence = {"date": date, "expense": expense, "amount": amount}
    print("\n--- Recorded Data ---")
    print(dailyexpence)
    
    save_expense_to_csv(date, expense, amount)
    totalexpence += amount_numeric
    expense_history.append(dailyexpence)
    
    if expense not in expense_category_history:
        expense_category_history.append(expense)
        expence_history.append(expense)
        expence_date_history.append(date)

    print("Current Total Expense:", totalexpence)
    print("-" * 40)
    return "success"

def show_monthly_history():
    month = input("Enter the month (1-12): ")
    print(f"\n--- Expense History for Month {month} ---")
    for record in expense_history:
        try:
            record_date = datetime.datetime.strptime(record["date"].split()[0], "%Y-%m-%d")
            if record_date.month == int(month):
                print(record)
        except Exception:
            print(f"Skipping record due to bad date format: {record}")

def show_yearly_history():
    year = input("Enter the year (YYYY): ")
    print(f"\n--- Expense History for Year {year} ---")
    for record in expense_history:
        try:
            record_date = datetime.datetime.strptime(record["date"].split()[0], "%Y-%m-%d")
            if record_date.year == int(year):
                print(record)
        except Exception:
            print(f"Skipping record due to bad date format: {record}")

def show_summary():
    print("\n================ FINAL SUMMARY ================")
    print("Total expense is:", totalexpence)
    print("The expense history is:", expense_history)
    print("The expense categories are:", expense_category_history)

print("Calibrating microphone... Please wait.")
with re.Microphone() as source:
    recognizer.adjust_for_ambient_noise(source, duration=1)

choice = int(input("do u want to hear the text (1 for Yes, 0 for No): "))

while True:
    status = add_expense(choice)
    if status == "stop":
        if choice == 1: speaker.Speak("Signing off.")
        break
    elif status == "continue":
        continue

show_summary()

yn = input("do u want to see the history? (y / n): ").lower()
if yn == "y":
    monthly_yn = input("do u want to see the monthly history? (y / n): ").lower()
    if monthly_yn == "y":
        show_monthly_history()

    yearly_yn = input("do u want to see the yearly history? (y / n): ").lower()
    if yearly_yn == "y":
        show_yearly_history()
            
    if monthly_yn != "y" and yearly_yn != "y":
        print("\n--- Full Expense History ---")
        for record in expense_history:
            print(record)
else:
    print("Exiting tracker system cleanly.")
