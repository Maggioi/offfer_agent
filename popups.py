import tkinter as tk
from tkinter import messagebox
import math

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
    popup.attributes("-topmost", True)
    popup.grab_set()

    # --- Liczba kolumn: max 10 opcji na kolumnę ---
    n_items = len(items)
    items_per_column = 10
    n_columns = max(1, math.ceil(n_items / items_per_column))

    # --- Nagłówek ---
    instrukcja = "Wybierz opcje (minimum jedna):" if wielokrotny else "Wybierz jedną opcję (wymagane):"
    tk.Label(popup, text=instrukcja, font=("Arial", 11, "bold")).pack(pady=15)

    # --- Przycisk na dole - pakowany PRZED items_frame ze stroną "bottom",       ---
    # --- dzięki czemu zawsze zostaje zarezerwowane dla niego miejsce i jest widoczny ---
    wynik = [] if wielokrotny else None
    zmienne_check = {}
    zmienna_radio = tk.StringVar(value="")

    def on_submit():
        nonlocal wynik

        if wielokrotny:
            wybrane = [nazwa for nazwa, var in zmienne_check.items() if var.get()]
            if not wybrane:
                messagebox.showwarning("Wymagany wybór", "Musisz zaznaczyć przynajmniej jedną opcję!")
                return
            wynik = wybrane
        else:
            wybrane = zmienna_radio.get()
            if not wybrane:
                messagebox.showwarning("Wymagany wybór", "Musisz zaznaczyć jedną opcję przed przesłaniem!")
                return
            wynik = wybrane

        popup.destroy()
        root.quit()

    submit_btn = tk.Button(
        popup, text="Zatwierdź", command=on_submit,
        bg="#4CAF50", fg="white", width=15
    )
    submit_btn.pack(side="bottom", pady=20)

    # --- Kontener na opcje (pakowany na środku, wypełnia resztę miejsca) ---
    items_frame = tk.Frame(popup)
    items_frame.pack(pady=5, fill="both", expand=True)

    column_frames = []
    for c in range(n_columns):
        col_frame = tk.Frame(items_frame)
        col_frame.grid(row=0, column=c, sticky="n", padx=15)
        column_frames.append(col_frame)

    wraplength = 340 if n_columns == 1 else 190

    for idx, item in enumerate(items):
        col_idx = idx // items_per_column
        target_frame = column_frames[col_idx]

        if wielokrotny:
            var = tk.BooleanVar(value=False)
            zmienne_check[item] = var
            btn = tk.Checkbutton(
                target_frame,
                text=item,
                variable=var,
                font=("Arial", 10),
                wraplength=wraplength,
                justify="left"
            )
        else:
            btn = tk.Radiobutton(
                target_frame,
                text=item,
                variable=zmienna_radio,
                value=item,
                font=("Arial", 10),
                wraplength=wraplength,
                justify="left"
            )
        btn.pack(anchor="w", pady=4)

    # --- Dopasowanie rozmiaru okna do RZECZYWISTEJ zawartości ---
    popup.update_idletasks()  # wymusza przeliczenie geometrii wszystkich widgetów

    req_width = popup.winfo_reqwidth() + 40   # margines bezpieczeństwa
    req_height = popup.winfo_reqheight() + 30  # margines bezpieczeństwa

    window_width = max(400, min(req_width, 1200))
    window_height = max(250, min(req_height, 750))

    # Wyśrodkowanie okna na ekranie
    screen_w = popup.winfo_screenwidth()
    screen_h = popup.winfo_screenheight()
    pos_x = (screen_w - window_width) // 2
    pos_y = (screen_h - window_height) // 2
    popup.geometry(f"{window_width}x{window_height}+{pos_x}+{pos_y}")

    root.mainloop()
    return wynik

def wybor_i_tekst_popup(items, etykieta_pola="Podaj wartość:", pytanie="Wybierz jedną opcję (wymagane):", wymagany_tekst=True):
    """Returns tuple(wynik)"""

    root = tk.Tk()
    root.withdraw()

    popup = tk.Toplevel()
    popup.title("Wybór i dane")
    popup.attributes("-topmost", True)
    popup.grab_set()

    # --- Liczba kolumn dla listy wyboru: max 10 opcji na kolumnę ---
    n_items = len(items)
    items_per_column = 10
    n_columns = max(1, math.ceil(n_items / items_per_column))

    # --- Nagłówek ---
    tk.Label(popup, text=pytanie, font=("Arial", 11, "bold")).pack(pady=(15, 5))

    wynik = [None, None]  # [wybrana_opcja, wpisany_tekst]
    zmienna_radio = tk.StringVar(value="")

    def on_submit():
        wybrana = zmienna_radio.get()
        if not wybrana:
            messagebox.showwarning("Wymagany wybór", "Musisz zaznaczyć jedną opcję przed przesłaniem!")
            return

        tekst = entry.get().strip()
        if wymagany_tekst and not tekst:
            messagebox.showwarning("Wymagane pole", "Pole tekstowe nie może być puste!")
            return

        wynik[0] = wybrana
        wynik[1] = tekst

        popup.destroy()
        root.quit()

    submit_btn = tk.Button(
        popup, text="Zatwierdź", command=on_submit,
        bg="#4CAF50", fg="white", width=15
    )
    submit_btn.pack(side="bottom", pady=20)

    # --- Pole tekstowe (nad przyciskiem, pod listą) ---
    text_frame = tk.Frame(popup)
    text_frame.pack(side="bottom", pady=(5, 10), padx=20, fill="x")

    tk.Label(text_frame, text=etykieta_pola, font=("Arial", 10)).pack(anchor="w")
    entry = tk.Entry(text_frame, font=("Arial", 11), width=30)
    entry.pack(fill="x", pady=(3, 0))
    entry.bind("<Return>", lambda event: on_submit())

    # --- Kontener na opcje wyboru ---
    items_frame = tk.Frame(popup)
    items_frame.pack(pady=5, fill="both", expand=True)

    column_frames = []
    for c in range(n_columns):
        col_frame = tk.Frame(items_frame)
        col_frame.grid(row=0, column=c, sticky="n", padx=15)
        column_frames.append(col_frame)

    wraplength = 340 if n_columns == 1 else 190

    for idx, item in enumerate(items):
        col_idx = idx // items_per_column
        target_frame = column_frames[col_idx]

        btn = tk.Radiobutton(
            target_frame,
            text=item,
            variable=zmienna_radio,
            value=item,
            font=("Arial", 10),
            wraplength=wraplength,
            justify="left"
        )
        btn.pack(anchor="w", pady=4)

    # --- Dopasowanie rozmiaru okna do RZECZYWISTEJ zawartości ---
    popup.update_idletasks()

    req_width = popup.winfo_reqwidth() + 40
    req_height = popup.winfo_reqheight() + 30

    window_width = max(400, min(req_width, 1200))
    window_height = max(300, min(req_height, 800))

    screen_w = popup.winfo_screenwidth()
    screen_h = popup.winfo_screenheight()
    pos_x = (screen_w - window_width) // 2
    pos_y = (screen_h - window_height) // 2
    popup.geometry(f"{window_width}x{window_height}+{pos_x}+{pos_y}")

    root.mainloop()
    return tuple(wynik)

if __name__ == "__main__":
    ...