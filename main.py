import tkinter as tk
from mult import mult_game_ui  # Updated to use the correct function name
from divi import division_game_ui  # Updated to use the correct function name
from subtraction import subtraction_game_ui  # Updated to use the correct function name
from addi import addition_game_ui  # Correct function name

def start_game(game_function, student_name):
    game_function(student_name)

def main_ui():
    root = tk.Tk()
    root.title("Math Quiz Program")

    tk.Label(root, text="Welcome to the Math Quiz Program!", font=("Arial", 14)).pack(pady=10)
    
    tk.Label(root, text="Enter your name:").pack()
    name_entry = tk.Entry(root)
    name_entry.pack()

    def on_choice(choice):
        student_name = name_entry.get()
        if student_name.strip():
            if choice == 1:
                start_game(mult_game_ui, student_name)  # Updated to use mult_game_ui
            elif choice == 2:
                start_game(division_game_ui, student_name)  # Updated to use division_game_ui
            elif choice == 3:
                start_game(subtraction_game_ui, student_name)  # Updated to use subtraction_game_ui
            elif choice == 4:
                start_game(addition_game_ui, student_name)  # Correct function name
            elif choice == 5:
                root.destroy()  # Properly closes the program
        else:
            tk.Label(root, text="Please enter your name before selecting a game.", fg="red").pack()

    tk.Label(root, text="\nChoose a game:").pack()
    tk.Button(root, text="Multiplication", command=lambda: on_choice(1)).pack()
    tk.Button(root, text="Division", command=lambda: on_choice(2)).pack()
    tk.Button(root, text="Subtraction", command=lambda: on_choice(3)).pack()
    tk.Button(root, text="Addition", command=lambda: on_choice(4)).pack()
    tk.Button(root, text="Exit", command=lambda: on_choice(5)).pack()

    root.mainloop()

if __name__ == "__main__":
    main_ui()