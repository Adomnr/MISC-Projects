#include <iostream>
#include <stdlib.h>
#include <string>
using namespace std;


const int floors = 3;
const int rooms = 3;
const int Maxium_one_customer = 10;
const int Maximum_customer = 10;

bool a[floors][rooms];
bool c[Maxium_one_customer][Maximum_customer];
int a1[floors][rooms];

class Customer {
private:
	string Name;
	int RoomNo;
	string room_allotment;
public:
	Customer() {  //Default Constructor (Does Nothing) !Made for array usage.
		RoomNo = 0;
	}
	Customer(string name, int roomtype) { //Constructor with Name and roomtype for manual use.
		Name = name;
		RoomBooking(roomtype);
	}
	~Customer() { // Deconstructor (information of user being deleted (Used in UI stage). //Will later use for deletion in user profile.
		cout << "Customer " << Name << " Deleted" << endl;
	}
	void Name_input(string name) { //For Manual Inputting of Name in a default constructor Profile.
		Name = name;
	}
	void Roomtype_input(int roomtype) {
		RoomBooking(roomtype);
	}
	void display() { // Displaying User information with the booked and available room.
		cout << "customer name is: " << Name << endl;
		cout << "customer Room Number is: " << RoomNo << endl;
		cout << "customer Room Type is: " << room_allotment << endl;
		cout << "Array is : " << endl;
		for (int i = 0; i < floors; i++) {
			for (int j = 0; j < rooms; j++) {
				cout << "    " << a[i][j];
			}
			cout << endl;
		}
	}
	int Room_Number_Automation(int roomtype, int b) { //Assigning Room Number for each room in the Hotel.
		if (b == 0) {
			RoomNo = 0;
			return 0;
		}
		RoomNo = (roomtype * 100) + (b);
		a1[roomtype-1][b - 1] = RoomNo;
		return RoomNo;
	}
	
	void Delete_Booking(int roomno){ // Making Room Available from Being Booked Turning Array equavilent to '0'.
		for (int i = 0; i < floors; i++) {
			for (int j = 0; j < rooms; j++) {
				if (roomno == a1[i][j]) {
					a[i][j] = false;
					break;
				}
			}
		}
		cout << "Room NO: " << roomno << " is Deleted" << endl;
	}

	void RoomBooking2(int roomtype) { // Logic for Room Booking (Mainly used for shorting of code written).
		int b = room_check(roomtype);
		Room_Number_Automation(roomtype, b);
		if (b != 0) {
			a[roomtype - 1][b - 1] = 1;
			cout << "You have successfully booked a room." << endl;
			Roomtype(roomtype);
		}
		else {
			cout << "Rooms are Full" << endl;
		}
	}

	void RoomBooking(int roomtype) { //Assiging the Room according to input provided in the following three categories.
		if (roomtype == 1) {
			RoomBooking2(roomtype);
		}
		else {
			if (roomtype == 2) {
				RoomBooking2(roomtype);
			}
			else {
				if (roomtype == 3) {
					RoomBooking2(roomtype);
				}
				else {
					cout << "Invalid Input! Unknown Error !" << endl;
				}
			}
		}
	}

	int room_check2(int roomtype, int maximum_room) { //Logic for checking if the rooms are available or not in the following floor.
		for (int i = 0; i < rooms; i++) {
			if (a[roomtype - 1][i] != 1) {
				return i + 1;
			}
			else {
				if (a[roomtype - 1][maximum_room] != 0) {
					return 0;
				}
				else
					continue;
			}
		}
	}

	int room_check(int roomtype) { // Checking the rooms availability using the logic function above
		int maximum_room = 2;
		if (roomtype == 1) {
			int c = room_check2(roomtype, maximum_room);
			return c;
		}
		else {
			if (roomtype == 2) {
				int c = room_check2(roomtype, maximum_room);
				return c;
			}
			else {
				if (roomtype == 3) {
					int c = room_check2(roomtype, maximum_room);
					return c;
				}
				else {
					return 0;
				}
			}
		}
		}
	void Roomtype(int roomtype) { //used for the assignment of normal ,medium and luxury display.
		if (roomtype == 1) {
			room_allotment = "Normal";
		}
		else {
			if (roomtype == 2) {
				room_allotment = "Medium";
			}
			else {
				if (roomtype == 3)
					room_allotment = "Luxury";
				else
					cout << "Invalid input!" << endl;
			}
		}
	}
};

void array_zero() {
	for (int i = 0; i < floors; i++) {
		for (int j = 0; j < rooms; j++) {
			a[i][j] = 0;
		}
	}
	for (int i = 0; i < Maxium_one_customer; i++) {
		for (int j = 0; j < Maximum_customer; j++) {
			c[i][j] = 0;
		}
	}
}


Customer Customer_initialization(Customer Arr[][Maximum_customer]) {
	Customer C1;
	string Name;
	int RoomType;
	char x, y;
	cout << "Do You Want to Enter a new Customer: "; cin >> x;
	if (x == 'y') {
		cout << "Enter Name of The Customer: ";
		cin >> Name;
		cout << "1 --  Normal\n2 -- Medium\n3 -- Luxury" << endl;
		cout << "Enter Roomtype of the Customer: ";
		cin >> RoomType;
		for (int i = 0; i < Maxium_one_customer; i++) {
			for (int j = 0; j < Maximum_customer; i++) {
				if (c[i][j] == 0) {
					Arr[i][j].Name_input(Name);
					Arr[i][j].Roomtype_input(RoomType);
					cout << "Do you want to see the display?" << endl; cin >> y;
					if (x == 'y') {
						Arr[i][j].display();
					}
					return Arr[i][j];
				}
			}
		}
	}
	else {
		if (x == 'n') {
			return C1;
		}
		else {
			cout << "Invalid Input (Unknown Error!" << endl;
			return C1;
		}
	}
}

/*Customer Customer_Space_Check(Customer Arr[][Maximum_customer]) {
	for (int i = 0; i < Maxium_one_customer; i++) {
		for (int j = 0; j < Maximum_customer; i++) {
			if (c[i][j] == 0) {
				return Arr[i][j];
			}
			else
				return;
		}
	}
}
*/

void Display_options() {

}

int main() {
	array_zero();
	Customer Arr[Maxium_one_customer][Maximum_customer];
	while (1) {
		Customer_initialization(Arr);
	}
}