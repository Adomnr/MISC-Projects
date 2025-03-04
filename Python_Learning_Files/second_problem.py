class_marks = {
    "History" : [],
    "Maths" : [],
    "Language Arts" : [],
    "Social Studies" : []
}

class_average = {
    "History Average" : 0,
    "Maths Average" : 0,
    "Language Arts Average" : 0,
    "Social Studies Average" : 0
}

alphabet_grades = {
    "History Grade" : "",
    "Maths Grade" : "",
    "Language Arts Grades" : "",
    "Social Studies Grades" : ""
}

def input_marks(string):
    inputs = input("How many "+string+" exam scores will do you have?: ")
    loop_inputs = int(inputs)
    while loop_inputs != 0:
        entry = int(input(string+" marks: "))
        if entry < 0 or entry >100:
            print("Invalid number, make sure your mark is between 0-100!")
            loop_inputs += 1
        else:
            class_marks[string].append(entry)
        loop_inputs -= 1
    return inputs

history_students = input_marks("History")
maths_students = input_marks("Maths")
language_arts_students = input_marks("Language Arts")
social_studies_students = input_marks("Social Studies")

def average_calculations(item):
    total = 0
    for x in item:
        x = int(x)
        total += x
    return total

class_average["History Average"] = average_calculations(class_marks.get("History"))/int(history_students)
class_average["Maths Average"] = average_calculations(class_marks.get("Maths"))/int(maths_students)
class_average["Language Arts Average"] = average_calculations(class_marks.get("Language Arts"))/int(language_arts_students)
class_average["Social Studies Average"] = average_calculations(class_marks.get("Social Studies"))/int(social_studies_students)

def grade_asssignment(x):
    x = int(x)
    if x > 90 and x <= 100:
        return "A"
    elif x > 75 and x <= 90:
        return "B"
    elif x > 60 and x<=75:
        return "C"
    elif x > 0  and x <= 60:
        return "D"
    else:
        print("Invalid Score")

alphabet_grades["History Grade"] = grade_asssignment(class_average["History Average"])
alphabet_grades["Maths Grade"] = grade_asssignment(class_average["Maths Average"])
alphabet_grades["Language Arts Grades"] = grade_asssignment(class_average["Language Arts Average"])
alphabet_grades["Social Studies Grades"] = grade_asssignment(class_average["Social Studies Average"])

print("\n",class_marks)
print("\n",class_average)
print("\n",alphabet_grades)