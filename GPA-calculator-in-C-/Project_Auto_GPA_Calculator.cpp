#include <iostream>
#include <stdlib.h>
#include <sstream>
using namespace std;

#define Size 80

float marks_calculation(float marks);

float display_gpa_full(float tot_cores, int tot_c_hours);

void display_gpa(string sub[], float cores[], float gpa[], int c_hours[], float tot_cores, float tot_c_hours, int n);

void gpa_s();

void welcome();

int main() {
	welcome();
}


float marks_calculation(float marks){
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

float display_gpa_full(float tot_cores, int tot_c_hours) {
	float result = 0;
	result = tot_cores / tot_c_hours;
	return result;
}

void display_gpa(string sub[], float cores[], float gpa[], int c_hours[] ,float tot_cores, float tot_c_hours, int n) {
	float gpa_d = 0;
	cout << "SN\tCH\tGPA\tCor" << endl;
	for (int i = 0; i < n; i++) {
		cout << sub[i] << "\t" << c_hours[i] << "\t" << gpa[i] << "\t" << cores[i] << endl;
	}
	cout << "Total corres is: " << tot_cores << endl;
	gpa_d = display_gpa_full(tot_cores, tot_c_hours);
	cout << "Full Gpa is: " << gpa_d << endl;

}

void gpa_s() {
	int n;
	int marks[Size];
	int tot_marks = 0;
	int c_hour[Size];
	int tot_c_hours = 0;
	string s_name[Size];
	float gpa[Size];
	float cores[Size];
	float tot_cores = 0;
	cout << "Enter number of subjects: ";
	cin >> n;
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
	display_gpa(s_name, cores, gpa, c_hour, tot_cores,tot_c_hours ,n);
}

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
			cout << "Later";
		else
			cout << "Invalid answer";
		return;
	}
}