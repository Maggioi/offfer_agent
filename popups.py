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

import tkinter as tk
from tkinter import messagebox

def wybor_popup(items, wielokrotny=False):
    root = tk.Tk()
    root.withdraw() 

    popup = tk.Toplevel()
    popup.title("Wybór materiału")
    popup.geometry("450x420")  # Delikatnie powiększone na wypadek wielu opcji
    popup.attributes("-topmost", True)
    popup.grab_set()

    # Dynamiczne dostosowanie nagłówka
    instrukcja = "Wybierz opcje (minimum jedna):" if wielokrotny else "Wybierz jedną opcję (wymagane):"
    tk.Label(popup, text=instrukcja, font=("Arial", 11, "bold")).pack(pady=15)

    # Przygotowanie struktur danych w zależności od trybu
    zmienne_check = {}  # Słownik na zmienne dla Checkbuttonów {nazwa: BooleanVar}
    zmienna_radio = tk.StringVar(value="")  # Pojedyncza zmienna dla Radiobuttonów

    # Generowanie przycisków
    for item in items:
        if wielokrotny:
            # Każdy checkbutton potrzebuje własnej zmiennej True/False
            var = tk.BooleanVar(value=False)
            zmienne_check[item] = var
            btn = tk.Checkbutton(
                popup, 
                text=item, 
                variable=var, 
                font=("Arial", 10),
                wraplength=380,   
                justify="left"
            )
        else:
            # Wszystkie radiobuttony dzielą jedną zmienną tekstową
            btn = tk.Radiobutton(
                popup, 
                text=item, 
                variable=zmienna_radio, 
                value=item, 
                font=("Arial", 10),
                wraplength=380,   
                justify="left"
            )
        btn.pack(anchor="w", padx=30, pady=4)

    # Zmienna na wynik (unikanie 'global')
    wynik = [] if wielokrotny else None

    def on_submit():
        nonlocal wynik  # Pozwala zmodyfikować zmienną 'wynik' z wyższego zakresu
        
        if wielokrotny:
            # Wyciągamy tylko te elementy, które zostały zaznaczone (True)
            wybrane = [nazwa for nazwa, var in zmienne_check.items() if var.get()]
            
            if not wybrane:
                messagebox.showwarning("Wymagany wybór", "Musisz zaznaczyć przynajmniej jedną opcję!")
                return  # Przerywa funkcję, nie zamyka okna
            
            wynik = wybrane
        else:
            wybrane = zmienna_radio.get()
            
            if not wybrane:
                messagebox.showwarning("Wymagany wybór", "Musisz zaznaczyć jedną opcję przed przesłaniem!")
                return  # Przerywa funkcję, nie zamyka okna
            
            wynik = wybrane

        popup.destroy()
        root.quit()
        
    submit_btn = tk.Button(popup, text="Zatwierdź", command=on_submit, bg="#4CAF50", fg="white", width=15)
    submit_btn.pack(pady=20)

    root.mainloop()
    return wynik

if __name__ == "__main__":
    ...