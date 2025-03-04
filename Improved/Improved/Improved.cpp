#include <iostream>
#include <stdlib.h>
#include <sstream>
using namespace std;

#define Size 134

float marks_calculation(int marks);

float display_gpa_full(float tot_cores, int tot_c_hours);

void display_gpa(string sub[], float cores[], float gpa[], int c_hours[], float tot_cores, float tot_c_hours, int n);

void gpa_s();

void welcome();

int main() {
	while (1) {
		welcome();
		cout << endl;
	}
}

float marks_calculation(int marks) {
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

float display_gpa_full(float tot_cores, float tot_c_hours) {
	float result = 0;
	result = tot_cores / tot_c_hours;
	return result;
}

void display_gpa(string sub[], float cores[], float gpa[], int c_hours[], float tot_cores, float tot_c_hours, int n) {
	float gpa_d = 0;
	cout << "SN\tCH\tGPA\tCor" << endl;
	for (int i = 0; i < n; i++) {
		cout << sub[i] << "\t" << c_hours[i] << "\t" << gpa[i] << "\t" << cores[i] << endl;
	}
	cout << "Total corres is: " << tot_cores << endl;
	gpa_d = display_gpa_full(tot_cores, tot_c_hours);
	cout << "Full Gpa is: " << gpa_d << endl;

}

void data_input(string s_name[],int marks[], int c_hour[],int& tot_c_hours, int i) {
		cout << "\nEnter subject name: "; cin >> s_name[i-1];
		cout << "\nEnter marks: "; cin >> marks[i-1];
		cout << "\nEnter credit hours: "; cin >> c_hour[i-1];
		tot_c_hours += c_hour[i-1];
		cout << endl;
}

void data_calculation(float gpa[], float cores[], int& tot_marks, float& tot_cores, int marks[], int c_hour[],int i) {
		gpa[i] = marks_calculation(marks[i]);
		cores[i] = gpa[i] * c_hour[i];
		tot_marks += marks[i];
		tot_cores += cores[i];
}

void gpa_s() {
	cout << "Choose your semester\n1)1-3\n2)4-6\n3)7-8" << endl;
	int semester;
	int n;
	cout << "Enter: "; cin >> semester;
	switch (semester) {
	case 1: n = 6; break;
	case 2: n = 5; break;
	case 3: n = 4; break;
	default: n = 0; cout << "Invalid input" << endl;
	}
	int s;
	char d;
	int marks[Size];
	int tot_marks = 0;
	int c_hour[Size];
	int tot_c_hours = 0;
	string s_name[Size];
	float gpa[Size];
	float cores[Size];
	float tot_cores = 0;
	for (int i = 0; i < n; i++) {
		for (int j = 0; j < n; j++) {
			cout << "Enter Subject " << j + 1 << endl;
		}
		cout << "Do you want to Display? Press(y/n) :"; cin >> d;
		if (d == 'y') {
			display_gpa(s_name, cores, gpa, c_hour, tot_cores, tot_c_hours, n);
		}
		cout << "Enter Subject: ";
		cin >> s;
		data_input(s_name, marks, c_hour,tot_c_hours,s);
		data_calculation(gpa, cores, tot_marks, tot_cores, marks, c_hour, s);
	}
	display_gpa(s_name, cores, gpa, c_hour, tot_cores, tot_c_hours, n);
}

/*void cgpa_s() {
	cout << "Enter your semester: ";
	int semester;
	cin >> semester;
	int n;
	switch (semester) {
	case 1: n = 18; break;
	case 2: n = 36; break;
	case 3: n = 54; break;
	case 4: n = 71; break;
	case 5: n = 88; break;
	case 6: n = 105; break;
	case 7: n = 120; break;
	case 8: n = 134; break;
	default: n = 0;  cout << "Invalid input Try again!" << endl;
	}
	int marks[Size];
	int tot_marks = 0;
	int c_hour[Size];
	int tot_c_hours = 0;
	string s_name[Size];
	float gpa[Size];
	float cores[Size];
	float tot_cores = 0;
	
	for (int i = 0; i < n; i++) {
		cout << "\nEnter subject name: "; cin >> s_name[i];
		cout << "\nEnter marks: "; cin >> marks[i];
		cout << "\nEnter credit hours: "; cin >> c_hour[i];
		tot_c_hours += c_hour[i];
	}
	for (int i = 0; i < n; i++) {
		gpa[i] = marks_calculation(marks[i]);
	}
	for (int i = 0; i < n; i++) {
		cores[i] = gpa[i] * c_hour[i];
	}
	for (int i = 0; i < n; i++) {
		tot_marks = marks[i];
	}
	for (int i = 0; i < n; i++) {
		tot_cores += cores[i];
	}
	display_gpa(s_name, cores, gpa, c_hour, tot_cores, tot_c_hours, n);
}
*/

void welcome() {
	cout << "Enter 1 for GPA \nEnter 2 for CGPA" << endl;
	cout << "What do you want to calculate: ";
	int n;
	cin >> n;
	if (n == 1) {
		gpa_s();
	}
	else {
		if (n == 2)
			cout << "Will be implemented with profiles later." << endl;//cgpa_s();
		else
			cout << "Invalid input try again." << endl;
		return;
	}
}