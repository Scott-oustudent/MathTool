from mult import mult_game
from divi import division_game
from subtraction import subtraction_game
from addi import addition_game

def main():
    print("Welcome to the Math Quiz Program!")
    student_name = input("Enter your name: ")

    while True:
        print("\nChoose a game:")
        print("1. Multiplication")
        print("2. Division")
        print("3. Subtraction")
        print("4. Addition")
        print("5. Exit")

        choice = input("Enter the number of your choice (1-5): ")
        if choice == "1":
            mult_game(student_name)
        elif choice == "2":
            division_game(student_name)
        elif choice == "3":
            subtraction_game(student_name)
        elif choice == "4":
            addition_game(student_name)
        elif choice == "5":
            print("Thank you for playing! Goodbye!")
            break
        else:
            print("Invalid choice. Please enter a number between 1 and 5.")

if __name__ == "__main__":
    main()
