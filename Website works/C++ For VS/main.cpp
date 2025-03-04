#include <iostream>
#include <stdlib.h>
#include <stdio.h>
using namespace std;

class Node {
public:
	int data;
	Node* next;

};

Node* head;

void Insert(int x);
void print();

int main() {
	head = NULL;
	cout << "How many numbers: "; int x, n; cin >> n;
	for (int i = 0; i < n; i++) {
		cout << "Enter the numbers: "; cin >> x;
		Insert(x);
		print();
	}
	cout << endl;
	system("PAUSE");
	return 0;
}

void Insert(int x) {
	Node* temp = new Node();
	temp->data = x;
	temp->next = head;
	head = temp;
}

void print() {
	Node* temp = head;
	cout << "list is: ";
	while (temp != NULL) {
		cout << temp->data << " ";
		temp = temp->next;
	}
}