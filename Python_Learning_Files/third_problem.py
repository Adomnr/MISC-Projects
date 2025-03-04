#Initializing global list
even_list = []
odd_list = []
all_list = []

#Defining a function that insert them into their respective library
def input_Number(x):
    #Exit Condition when "quit" is inserted
    if x =="quit":
        return
    x = int(x)
    #Deviates Negative Numbers
    if x <= 0:
        print("Make sure your number is above 0!")
        welcome_fc() #F
    #Put Positive integer to verify which one them qualifies for each list.
    else:
        even_numbers(x) #F
        odd_numbers(x) #F
        all_numbers(x) #F

#Funtion to identify even nunmbers
def even_numbers(x):
    if x % 2 == 0:
        even_list.append(x) #F

#Function to identify odd numbers
def odd_numbers(x):
    if x % 2 != 0:
        odd_list.append(x) #F

#Function for appending every number into all list.
def all_numbers(x):
    all_list.append(x) #F

#Function that prints output required by the problem
def output():
    #Printing Every Number inside the list.
    print("All List: ", all_list)
    print("Even List: ", even_list)
    print("Odd List: ", odd_list)

    #Prints maximum and Minimum NUmber of each list.
    if len(all_list) > 0:
        print("\nMinimum Number: ", min(all_list), "Maximum Number: ", max(all_list))
    else:
        print("\nNo numbers were entered")
    if len(even_list) > 0:
        print("\nMinimum Even Number: ", min(even_list), "Maximum Even Number: ", max(even_list))
    else:
        print("\nNo even numbers entered")
    if len(odd_list) > 0:
        print("\nMinimum Odd Number: ", min(odd_list), "Maximum Odd Number: ", max(odd_list))
    else:
        print("\nNo odd numbers entered")

#Function to take input from user and calls printing function
def welcome_fc():
    while True:
        #Input from User.
        given = input("Enter a positive Number/Type quit to stop:" ).lower()
        if given == "quit":
            break
        else:
            input_Number(given)
    output() #F

welcome_fc() #F