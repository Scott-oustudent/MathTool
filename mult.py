import random
from openpyxl import Workbook, load_workbook
from datetime import datetime

def save_to_excel(student_name, game_type, correct_ans, wrong_ans, questions):
    file_name = "students_scores.xlsx"
    try:
        workbook = load_workbook(file_name)
        sheet = workbook.active
    except FileNotFoundError:
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Date", "Student Name", "Game Type", "Correct Answers", "Wrong Answers", "Questions Asked"])

    questions_str = "; ".join(questions)
    date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    sheet.append([date, student_name, game_type, correct_ans, wrong_ans, questions_str])
    workbook.save(file_name)

def mult_game(student_name):
    correct_ans = 0
    wrong_ans = 0
    questions = []
    asked_questions = set()  # Set to track unique questions

    work = int(input('What Multiplication Table are we working on today? \n(Please choose from 1 to 10)\nWrite your answer here: '))

    counter = 0
    while counter < 10:
        num = random.randint(1, 10)
        question = f"{work} x {num}"

        # Ensure the question is unique
        if question in asked_questions:
            continue
        asked_questions.add(question)

        ans = num * work
        try:
            stu_ans = int(input(f"What is {question}: "))
            if ans == stu_ans:
                correct_ans += 1
                print("Correct. Well done!!!")
            else:
                wrong_ans += 1
                print("No sorry. This is incorrect!")
            questions.append(f"{question} = {stu_ans} (Correct: {stu_ans == ans})")
            counter += 1
        except ValueError:
            print("Please enter a valid number.")
            questions.append(f"{question} = Invalid Input")

    print("\nSummary:")
    print(f"Correct Answers: {correct_ans}")
    print(f"Wrong Answers: {wrong_ans}")

    save_to_excel(student_name, "Multiplication", correct_ans, wrong_ans, questions)
