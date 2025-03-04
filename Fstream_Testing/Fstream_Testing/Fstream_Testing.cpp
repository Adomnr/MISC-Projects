#include <iostream>
#include <stdlib.h>
#include <stdio.h>
#include <string>
#include <fstream>
using namespace std;

int main() {
	fstream Departments;
	int input;
	int counting = 1;
	Departments.open("Departments,txt", ios::out);
	Departments << "Electrical_Enginnering" << endl;
	Departments << "Applied Sciences" << endl;
	Departments.close();
	Departments.open("Departments,txt", ios::in);
	if (Departments.is_open()) {
		string line;
		while (getline(Departments, line)) {
			cout << counting << "-" << line << endl;
			counting++;
		}
	}
	Departments.close();
	cout << "Enter Department: "; cin >> input;
	switch (input) {
	case 1:
		counting = 1;
		Departments.open("Departments,txt", ios::in);
		if (Departments.is_open()) {
			string line;
			while (getline(Departments, line)) {
				if (counting == input) {
					cout << " Welcome to " << line << " Department!" << endl;
				}
				counting++;
			}
		}
		cout << endl;
		break;
	case 2:
		counting = 1;
		if (Departments.is_open()) {
			string line;
			while (getline(Departments, line)) {
				if (counting == input) {
					cout << " Welcome to " << line << " Department!" << endl;
				}
				counting++;
			}
		}
		cout << endl;
		break;
	default:
		cout << "Invalid" << endl;
	}
}