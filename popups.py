import tkinter as tk
from tkinter import messagebox

def show_yes_no_popup():
    # Hide the main root window so only the dialog box appears
    root = tk.Tk()
    root.withdraw()
    
    # Make sure the dialog handles layering and shows up on top
    root.attributes("-topmost", True)

    # Trigger the native system Yes/No popup
    # Arguments: (Window Title, Prompt Message)
    user_choice = messagebox.askyesno("System Confirmation", "Czy liczysz OP50?")

    # Process the result
    if user_choice:
        print("User clicked YES.")
        # Put your 'Yes' logic here
    else:
        print("User clicked NO or closed the window.")
        # Put your 'No' logic here

    root.destroy()

def wybor_plyty_popup():
    root = tk.Tk()
    root.withdraw() 

    popup = tk.Toplevel()
    popup.title("Wybór materiału")
    popup.geometry("450x380")
    popup.attributes("-topmost", True)
    popup.grab_set()

    final_choice = "AAAAAAAAAA"
    items = [
        "CPL on M3 class chipboard (D-s2,d0)", 
        "CPL on M1 class (s/fire) chipboard (B-s2,d0)", 
        "HPL on M3 class chipboard (D-s2,d0)", 
        "HPL on M1 class (s/fire) chipboard (B-s2,d0)",
        "Veneer on M3 chipboard (D-s2,d0)", 
        "Veneer on M1 (s/fire) chipboard (B-s2,d0)"
    ]
    
    # Jedna wspólna zmienna dla wszystkich opcji. 
    # Ustawiamy wartość początkową na pusty string, żeby domyślnie nic nie było zaznaczone.
    selected_item = tk.StringVar(value="")

    tk.Label(popup, text="Wybierz jedną opcję (wymagane):", font=("Arial", 11, "bold")).pack(pady=15)

    # Generowanie przycisków Radiobutton
    for item in items:
        rb = tk.Radiobutton(
            popup, 
            text=item, 
            variable=selected_item,  # Wszystkie przyciski współdzielą tę samą zmienną
            value=item,              # Tę wartość przyjmie zmienna, gdy klikniesz ten przycisk
            font=("Arial", 10),
            wraplength=380,   
            justify="left"    
        )
        rb.pack(anchor="w", padx=30, pady=4)

    def on_submit():
        global choice
        # Pobieramy aktualnie zaznaczoną wartość
        choice = selected_item.get()
        
        # Walidacja: Jeśli zmienna jest pusta, użytkownik nic nie kliknął
        if not choice:
            messagebox.showwarning(
                "Wymagany wybór", 
                "Musisz zaznaczyć jedną opcję przed przesłaniem!"
            )
        
       # messagebox.showinfo("Sukces", f"Wybrano:\n\n{choice}")
        popup.destroy()
        root.quit()
        
    
    submit_btn = tk.Button(popup, text="Zatwierdź", command=on_submit, bg="#4CAF50", fg="white", width=15)
    submit_btn.pack(pady=20)

    root.mainloop()
    return choice

if __name__ == "__main__":
    show_yes_no_popup()
    wybor_plyty_popup()