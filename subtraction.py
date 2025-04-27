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
        # Add headers, including a column for questions
        sheet.append(["Date", "Student Name", "Game Type", "Correct Answers", "Wrong Answers", "Questions Asked"])

    # Format the questions as a single string
    questions_str = "; ".join(questions)

    # Add a new row with the student's data
    date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    sheet.append([date, student_name, game_type, correct_ans, wrong_ans, questions_str])
    workbook.save(file_name)

def subtraction_game(student_name):
    correct_ans = 0
    wrong_ans = 0
    questions = []  # List to store the questions asked

    print("\nWelcome to the Subtraction Quiz!")
    counter = 0
    while counter < 10:  # Loop for 10 questions
        num1 = random.randint(1, 50)
        num2 = random.randint(1, 50)
        if num1 < num2:
            num1, num2 = num2, num1  # Ensure no negative results
        answer = num1 - num2

        question = f"{num1} - {num2}"  # Format the question
        try:
            stu_ans = int(input(f"What is {question}: "))
            if stu_ans == answer:
                correct_ans += 1
                print("Correct. Well done!!!")
            else:
                wrong_ans += 1
                print(f"No sorry. The correct answer was {answer}")
            questions.append(f"{question} = {stu_ans} (Correct: {stu_ans == answer})")  # Log the question and result
            counter += 1
        except ValueError:
            print("Please enter a valid number.")
            questions.append(f"{question} = Invalid Input")  # Log invalid input

    print("\nSubtraction Quiz Summary:")
    print(f"Correct Answers: {correct_ans}")
    print(f"Wrong Answers: {wrong_ans}")

    # Save the student's scores and questions to Excel
    save_to_excel(student_name, "Subtraction", correct_ans, wrong_ans, questions)

def main():
    print("Welcome to the Math Quiz Program!")
    student_name = input("Enter your name: ")

    print("Choose a game:")
    print("1. Multiplication")
    print("2. Division")
    print("3. Subtraction")
    print("4. Addition")

    choice = input("Enter the number of your choice: ")

    if choice == "1":
        mult_game(student_name)
    elif choice == "2":
        division_game(student_name)
    elif choice == "3":
        subtraction_game(student_name)
    elif choice == "4":
        addition_game(student_name)
    else:
        print("Invalid choice. Please restart the program.")

if __name__ == "__main__":
    main()
