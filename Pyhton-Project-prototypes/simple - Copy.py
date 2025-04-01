print("Enter your daily expenses :")

# Initialize the date variable
date = ''

# Loop until the user enters 'done'
while date != 'done':
    date = input("Enter the date or 'done' to finish: ")
    if date == 'done':
        break  # Exit the loop if 'done' is entered

    # Initialize the total expenses for the day
    total_expenses = 0
    expense = ''

    # Loop until the user enters 'done' for expenses
    while expense != 'done':
        expense = input("Enter expense amount or 'done' to finish: ")
        if expense == 'done':
            break  # Exit the loop if 'done' is entered

        # Convert the expense to a float and then to an integer
        expense = int(float(expense))
        
        # Add the expense to the total expenses for the day
        total_expenses += expense

    # Open the file in append mode
    file = open('expenses.txt', mode='a')
   
    # Write the date and total expenses for the day in file
    file.write(date + ": Rs." + str(total_expenses) + "\n")
   
    # Close the file
    file.close()
