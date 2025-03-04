
def room_calculation():
    (standard_room, deluxe_room, suite) = (150, 180, 220)
    option = input("Would you like to book a room? (Yes/No) ").lower()
    if option == "yes":
        print("\nWhat type of room do you want?\nWe have the following options:\n1 - Standard Room :  Cost $150 per night\n2 - Deluxe Room :  Cost $180 per night\n3 - Suite :  Cost $220 per night\nPlease select 1, 2, or 3")
        x = int(input("Enter from the previous options: "))
        cost = 0
        if x == 1:
            cost = standard_room
        elif x == 2:
            cost = deluxe_room
        elif x == 3:
            cost = suite
        else:
            print("Invalid Input! Please try again\n")
            room_calculation() #F
        stay_length = input("Enter length of stay (In days) : ")
        total_cost = int(cost) * int(stay_length)
        return total_cost
    elif option == "no":
        return False
    else:
        print("Invalid Input! Please try again\n")
        room_calculation() #F

def rental_cars():
    print("\nRental cars are $100 up to 24 hours then $60 for each additional day")
    x = input("Would you like to rent a car? (Yes/No) ").lower()
    rent_car_cost = 0
    if x == "yes":
        rent_days = int(input("\nEnter amount of days you will rent a car: "))
        if rent_days >= 1:
            rent_car_cost = 100 + ((rent_days - 1) * 60)
        else:
            print("Invalid Input! Please try again\n")
            rental_cars() #F
        return rent_car_cost
    elif x == "no":
        rent_car_cost = 0
    else:
        print("Invalid Input! Please try again\n")
        rental_cars() #F
    return rent_car_cost

def theme_park_ticket():
    print("\nTheme park parking prices are as follows:\n$175 for 1 day and $50 additional for each extended day")
    x = input("Would you like to buy parking? (Yes/No) ").lower()
    if x == "yes":
        ticket_length = int(input("Enter duration of parking (In days): "))
        ticket_cost = 0
        if ticket_length >= 1:
            ticket_cost = 175 + ((ticket_length - 1) * 50)
        else:
            print("Invalid Input! Please try again\n")
            theme_park_ticket() #F
        return ticket_cost
    elif x == "no":
        return False
    else:
        print("Invalid Input! Please try again\n")
        theme_park_ticket() #F

def theme_park_trip():
    x = int(room_calculation())
    y = int(rental_cars())
    z = int(theme_park_ticket())
    print("\nRoom Cost: ", x)
    print("Rental Car Cost: ", y)
    print("Theme Park Parking Cost: ", z)
    print("Total theme park trip cost will be: ", (x + y + z))

theme_park_trip() #F