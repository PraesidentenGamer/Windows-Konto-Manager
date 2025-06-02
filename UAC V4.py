import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import wmi
import subprocess
import ctypes
import sys

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def get_users_wmi():
    c = wmi.WMI()
    users = []
    for user in c.Win32_UserAccount(LocalAccount=True):
        name = user.Name
        status = "Aktiv" if not user.Disabled else "Inaktiv"
        users.append((name, status))
    return users

def run_net_command(args):
    try:
        completed = subprocess.run(["net"] + args, capture_output=True, text=True, shell=True)
        success = (completed.returncode == 0)
        # stdout oder stderr könnten None sein, daher absichern
        output = (completed.stdout or "") + (completed.stderr or "")
        return success, output
    except Exception as e:
        return False, str(e)

def refresh_users():
    for item in tree.get_children():
        tree.delete(item)
    for user, status in get_users_wmi():
        tree.insert("", "end", values=(user, status))

def disable_user():
    user = get_selected_user()
    if not user:
        return
    if messagebox.askyesno("Bestätigung", f"Benutzer '{user}' deaktivieren?"):
        success, output = run_net_command(["user", user, "/active:no"])
        if success:
            messagebox.showinfo("Erfolg", f"Benutzer '{user}' wurde deaktiviert.")
            refresh_users()
        else:
            messagebox.showerror("Fehler", f"Fehler beim Deaktivieren:\n{output}")

def enable_user():
    user = get_selected_user()
    if not user:
        return
    if messagebox.askyesno("Bestätigung", f"Benutzer '{user}' aktivieren?"):
        success, output = run_net_command(["user", user, "/active:yes"])
        if success:
            messagebox.showinfo("Erfolg", f"Benutzer '{user}' wurde aktiviert.")
            refresh_users()
        else:
            messagebox.showerror("Fehler", f"Fehler beim Aktivieren:\n{output}")

def delete_user():
    user = get_selected_user()
    if not user:
        return
    if messagebox.askyesno("Warnung", f"Benutzer '{user}' wirklich löschen?"):
        success, output = run_net_command(["user", user, "/delete"])
        if success:
            messagebox.showinfo("Erfolg", f"Benutzer '{user}' wurde gelöscht.")
            refresh_users()
        else:
            messagebox.showerror("Fehler", f"Fehler beim Löschen:\n{output}")

def set_password():
    user = get_selected_user()
    if not user:
        return
    new_pw = simpledialog.askstring("Passwort setzen", f"Neues Passwort für '{user}':", show="*")
    if new_pw is None:
        return
    if messagebox.askyesno("Bestätigung", f"Passwort für '{user}' ändern?"):
        success, output = run_net_command(["user", user, new_pw])
        if success:
            messagebox.showinfo("Erfolg", f"Passwort für '{user}' wurde geändert.")
        else:
            messagebox.showerror("Fehler", f"Fehler beim Ändern des Passworts:\n{output}")

def create_user():
    username = simpledialog.askstring("Neuen Benutzer erstellen", "Benutzername eingeben:")
    if not username:
        return
    
    password = simpledialog.askstring("Passwort setzen", f"Passwort für '{username}' (leer = kein Passwort):", show="*")
    
    # Gruppen-Auswahl (z.B. Administratoren oder Benutzer)
    groups = ["Administratoren", "Benutzer"]
    group = simpledialog.askstring("Gruppenwahl", f"Gruppe für '{username}' wählen:\n" + "\n".join(groups))
    if group not in groups:
        messagebox.showerror("Fehler", "Ungültige Gruppe ausgewählt.")
        return
    
    args = [username]
    if password:
        args.append(password)
    else:
        args.append("*")  # * = kein Passwort
    args.append("/add")
    
    if messagebox.askyesno("Bestätigung", f"Benutzer '{username}' mit Passwort {'gesetzt' if password else 'leer'} erstellen?"):
        success, output = run_net_command(["user"] + args)
        if success:
            # Gruppe hinzufügen
            add_group_success, add_group_output = run_net_command(["localgroup", group, username, "/add"])
            if add_group_success:
                messagebox.showinfo("Erfolg", f"Benutzer '{username}' wurde erstellt und zur Gruppe '{group}' hinzugefügt.")
            else:
                messagebox.showwarning("Warnung", f"Benutzer wurde erstellt, aber Gruppen-Zuweisung schlug fehl:\n{add_group_output}")
            refresh_users()
        else:
            messagebox.showerror("Fehler", f"Fehler beim Erstellen des Benutzers:\n{output}")

def get_selected_user():
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("Keine Auswahl", "Bitte zuerst einen Benutzer auswählen.")
        return None
    return tree.item(selected[0])["values"][0]

if not is_admin():
    messagebox.showerror("Adminrechte benötigt", "Bitte das Programm als Administrator starten.")
    sys.exit(1)

# === GUI Styling ===

root = tk.Tk()
root.title("Windows Benutzerkonten Verwalten")
root.geometry("600x480")
root.configure(bg="#2e2e2e")

style = ttk.Style(root)
style.theme_use('clam')
style.configure("Treeview",
                background="#383838",
                foreground="white",
                rowheight=28,
                fieldbackground="#383838",
                font=("Segoe UI", 11))
style.map('Treeview', background=[('selected', '#007acc')])

style.configure("TButton",
                background="#007acc",
                foreground="white",
                font=("Segoe UI", 11, "bold"),
                padding=6)
style.map("TButton",
          background=[('active', '#005f99')])

# Baumansicht
tree = ttk.Treeview(root, columns=("Benutzername", "Status"), show="headings")
tree.heading("Benutzername", text="Benutzername")
tree.heading("Status", text="Status")
tree.column("Benutzername", width=350, anchor="w")
tree.column("Status", width=150, anchor="center")
tree.pack(fill="both", expand=True, padx=15, pady=(15,5))

# Buttons mit Abstand in Frame
frame = tk.Frame(root, bg="#2e2e2e")
frame.pack(pady=10)

btn_refresh = ttk.Button(frame, text="Aktualisieren", command=refresh_users)
btn_refresh.grid(row=0, column=0, padx=6)

btn_create = ttk.Button(frame, text="Benutzer erstellen", command=create_user)
btn_create.grid(row=0, column=1, padx=6)

btn_enable = ttk.Button(frame, text="Aktivieren", command=enable_user)
btn_enable.grid(row=0, column=2, padx=6)

btn_disable = ttk.Button(frame, text="Deaktivieren", command=disable_user)
btn_disable.grid(row=0, column=3, padx=6)

btn_delete = ttk.Button(frame, text="Löschen", command=delete_user)
btn_delete.grid(row=0, column=4, padx=6)

btn_set_pw = ttk.Button(frame, text="Passwort setzen", command=set_password)
btn_set_pw.grid(row=0, column=5, padx=6)

# Fußzeile
footer = tk.Label(root, text="Hinweis: Administratorrechte erforderlich", bg="#2e2e2e", fg="#cccccc", font=("Segoe UI", 9))
footer.pack(side="bottom", pady=5)

refresh_users()

root.mainloop()
