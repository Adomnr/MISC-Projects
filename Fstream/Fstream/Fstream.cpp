#include <iostream>
#include <stdlib.h>
#include <stdio.h>
#include <string>
#include <fstream>
using namespace std;


float marks_calculation(int marks);
void department_output(int counting, int input, string Profile_Name);
void input_data();
void data_entry();
void condition(string equal);
void semester_data(string equal);
void semester_output(int counting, int x);
void semester_selection();
void welcome(string Profile_Name);
void create_profile(string& Profile_Name, string& Profile_Password);
void login_profile(string& Profile_Name, string& Profile_Password, string& Prof_Name);
void setup();


int main() {
	while (1) {
		cout << "Do you want to continue(y/n): ";
		char y;
		cin >> y;
		if (y == 'n') {
			break;
		}
		setup();
		cout << endl;
	}
}


float marks_calculation(int marks) { // Will later be used to calculated gpa with a slight modification.
	float result = 0.0;
	if (marks > 100)
		return false;
	if (marks < 0)
		return false;
	if (marks < 50) {
		return result;
	}
	else {
		if (marks >= 50 && marks <= 68) {

			result = 2 + ((marks - 50) / 18);
			return result;
		}
		else {
			if (marks >= 69 && marks <= 85) {
				result = 3 + ((marks - 68) / 17);
				return result;
			}
			else {
				result = 4;
				return result;
			}
		}
	}
}

void department_output(int counting, int input,string Profile_Name) { // Outputs department Welcome to the user.
	fstream Departments;
	Departments.open("Departments.txt", ios::in);
	if (Departments.is_open()) {
		string line;
		while (getline(Departments, line)) {
			if (counting == input) {
				cout << Profile_Name << " Welcome to " << line << " Department!" << endl;
			}
			counting++;
		}
	}
	cout << endl;
	Departments.close();
}

void input_data() { // Will be used to store marks by inputting.
	fstream profile;
	cout << "Will be implemented later" << endl;
}

void data_entry() { // Will be used to select semester to be enter data.
	char input;
	cout << "Do you want to Continue(y/n): "; cin >> input;
	if (input == 'y') {
		input_data();
	}
	else {
		if (input == 'n') {
			semester_selection();
		}
		else {
			cout << "Invalid Input!" << endl;
			data_entry();
		}
	}
}

void condition(string equal) { // Displays semester subjects to the user and confirmation using data entry function.
	fstream Semester;
	fstream Semester_data;
	Semester_data.open("Semester_data.txt", ios::in);
	Semester.open("Semester.txt", ios::in);
	int counting = 1;
	if (Semester.is_open()) {
		string line;
		while (getline(Semester, line)) {
			if (line == equal && counting <= 3) {
				if (Semester_data.is_open()) {
					while (getline(Semester_data, line)) {
						if (line == equal) {
							string line2;
							int internal_counting = 1;
							while (getline(Semester_data, line2)) {
								cout << line2 << endl;
								if (internal_counting > 6) {
									data_entry();
									break;
								}
								internal_counting++;
							}
						}
					}
				}
			}
			if (line == equal && counting >= 4 && counting <= 6) {
				if (Semester_data.is_open()) {
					while (getline(Semester_data, line)) {
						if (line == equal) {
							string line2;
							int internal_counting = 1;
							while (getline(Semester_data, line2)) {
								cout << line2 << endl;
								if (internal_counting > 5) {
									data_entry();
									break;
								}
								internal_counting++;
							}
						}
					}
				}
			}
			if (line == equal && counting >= 7 && counting <= 8) {
				if (Semester_data.is_open()) {
					while (getline(Semester_data, line)) {
						if (line == equal) {
							string line2;
							int internal_counting = 1;
							while (getline(Semester_data, line2)) {
								cout << line2 << endl;
								if (internal_counting > 4) {
									data_entry();
									break;
								}
								internal_counting++;
							}
						}
					}
				}
			}
			counting++;
		}
	}
	Semester.close();
	Semester_data.close();
	
}

void semester_data(string equal) {
	fstream Semester_data;
	Semester_data.open("Semester_data.txt", ios::app);
	Semester_data.close();
	Semester_data.open("Semester_data.txt", ios::in);
	string line;
	if (Semester_data.is_open()) {
		while (getline(Semester_data, line)) {
			if (line == equal) {
				condition(equal);
			}
		}
	}
	Semester_data.close();
}

void semester_output(int counting, int x) {
	fstream Semester;
	Semester.open("Semester.txt", ios::in);
	if (Semester.is_open()) {
		string line;
		while (getline(Semester, line)) {
			if (counting == x) {
				cout << line << endl;
				semester_data(line);
			}
			counting++;
		}
	}
	Semester.close();
}

void semester_selection() {
	int x;
	int counting = 1;
	cout << "Select your semester: " << endl;
	fstream Semester;
	Semester.open("Semester.txt", ios::in);
	if (Semester.is_open()) {
		string line;
		while (getline(Semester, line)) {
			cout << line << endl;
		}
	}
	Semester.close();
	cout << "Enter: "; cin >> x;
	switch (x) {
	case 1:
		semester_output(counting, x);
		break;
	case 2:
		semester_output(counting, x);
		break;
	case 3:
		semester_output(counting, x);
		break;
	case 4:
		semester_output(counting, x);
		break;
	case 5:
		semester_output(counting, x);
		break;
	case 6:
		semester_output(counting, x);
		break;
	case 7:
		semester_output(counting, x);
		break;
	case 8:
		semester_output(counting, x);
		break;
	default:
		cout << "Invlaid Input!" << endl;
	}
	Semester.close();
}

void welcome(string Profile_Name) {
	cout << "Enter your Department from Below: " << endl;
	fstream Departments;
	int counting = 1;
	int input;
	Departments.open("Departments.txt", ios::in);
	if (Departments.is_open()) {
		string line;
		while(getline(Departments, line)){
			cout << counting << "-" << line << endl;
			counting++;
		}
	}
	Departments.close();
	counting = 1;
	cout << "Enter Department: "; cin >> input;
	switch (input) {
	case 1:
		department_output(counting, input, Profile_Name);
		cout << endl;
		break;
	case 2:
		department_output(counting, input, Profile_Name);
		cout << endl;
		break;	
	case 3:
		department_output(counting, input, Profile_Name);
		cout << endl;
		break;
	default:
		cout << "Invalid Input!" << endl;
		welcome(Profile_Name);
	}
	cout << "Enter 1 for GPA \nEnter 2 for CGPA" << endl;
	cout << "What do you want to calculate: ";
	int n;
	cin >> n;
	if (n == 1) {
		semester_selection();
	}
	else {
		if (n == 2)
			cout << "Will be implemented with profiles later." << endl;//cgpa_s();
		else
			cout << "Invalid input try again." << endl;
		return;
	}
}

void create_profile(string& Profile_Name, string& Profile_Password) {  //Creating new files for new profiles.
	string Name;
	string Password;
	cout << "Enter Name: "; cin >> Name;
	Name += ".txt";
	cout << "Create Password: "; cin >> Password;
	fstream Profile;
	Profile.open(Name, ios::out);
	Profile.close();
	Profile_Name = Name;
	Profile_Password = Password;
}

void login_profile(string& Profile_Name, string& Profile_Password, string& Prof_Name) { //Takes UserName and Password from the user.
	string Name;
	string Password;
	cout << "Enter UserName: "; cin >> Name;
	cout << "Enter Password: "; cin >> Password;
	Prof_Name = Name;
	Name += ".txt";
	Profile_Name = Name;
	Profile_Password = Password;
}

void setup() { // Does all the checking and interfacing for the signing up and login of the profiles.
	string Name;
	fstream all_profiles;
	string Profile_Name;
	string Profile_Password;
	all_profiles.open("All_Profiles.txt", ios::app);
	all_profiles.close();
	cout << "New User? Sign UP\n1 - Sign Up\nExisting User?\n2 - Login" << endl;
	int x;
	cout << "Enter : "; cin >> x;
	if (x == 1) {
		cout << "Sign UP" << endl;
		create_profile(Profile_Name, Profile_Password);
		all_profiles.open("All_Profiles.txt", ios::app);
		if (all_profiles.is_open()) {
			all_profiles << Profile_Name +'\t'+Profile_Password << endl;
		}
		all_profiles.close();
	}
	else {
		if (x == 2) {
			cout << "Logn in" << endl;
			all_profiles.open("All_Profiles.txt", ios::in);
			login_profile(Profile_Name, Profile_Password, Name);
			if (all_profiles.is_open()) {
				string line;
				string equal = Profile_Name + '\t' + Profile_Password;
				while (getline(all_profiles, line)) {
					if (line == equal) {
						cout << endl;
						cout << "You have successfully logged in" << endl;
						welcome(Name);
					}
				}
				all_profiles.close();
			}
		}
		else {
			cout << "Invalid Input try again" << endl << endl;
			setup();
		}
	}
	all_profiles.close();
}
