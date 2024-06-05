import tkinter as tk

from pathlib import Path
from tkinter import filedialog, ttk
from subprocess import check_output


def reset_progress():
    style.configure("TProgressbar", background="blue", troughcolor="lightgrey")
    progress_bar["value"] = 0


def on_progress_complete():
    style.configure("TProgressbar", background="green", troughcolor="lightgrey")


def run_script():
    reset_progress()

    input_folder = input_folder_entry.get()
    output_folder = output_folder_entry.get()
    output_format = output_format_var.get()

    if not input_folder:
        tk.messagebox.showerror("Error", "Please select an input folder.")
        return
    input_folder = Path(input_folder)
    if not output_folder:
        tk.messagebox.showerror("Error", "Please select an output folder.")
        return
    output_folder = Path(output_folder)
    if not output_folder.exists():
        output_folder.mkdir(parents=True)

    runtime = Path('./BTSDAExReport.exe')
    if not runtime.exists():
        tk.messagebox.showerror("Error", "BTSDAExReport.exe not found.")
        return

    progress_bar["maximum"] = len(list(input_folder.glob("*.nda*")))

    # Simulate a long-running task
    for index, file in enumerate(input_folder.glob("*.nda*")):
        target_file = output_folder / f"{file.stem}.xlsx"
        cmd = f'"{runtime}" export {output_format} "{file}" "{target_file}"'
        try:
            check_output(cmd, shell=True)
        except Exception as e:
            tk.messagebox.showerror("Error", f"Error converting {file}: {e}")
            return
        progress_bar["value"] = index + 1
        root.update()

    on_progress_complete()


# The window
root = tk.Tk()
root.title("Converter GUI for Neware NDA files")

tk.Label(root, text="Input Folder:").grid(row=0, column=0)
input_folder_entry = tk.Entry(root)
input_folder_entry.grid(row=0, column=1)
tk.Button(
    root,
    text="Browse",
    command=lambda: input_folder_entry.delete(0, tk.END)
    or input_folder_entry.insert(0, filedialog.askdirectory()),
).grid(row=0, column=2)

tk.Label(root, text="Output Folder:").grid(row=1, column=0)
output_folder_entry = tk.Entry(root)
output_folder_entry.grid(row=1, column=1)
tk.Button(
    root,
    text="Browse",
    command=lambda: output_folder_entry.delete(0, tk.END)
    or output_folder_entry.insert(0, filedialog.askdirectory()),
).grid(row=1, column=2)

tk.Label(root, text="Output Format:").grid(row=2, column=0)
output_format_var = tk.StringVar()
tk.OptionMenu(root, output_format_var, "General", "Layer", "Custom", "Sim").grid(
    row=2, column=1
)
output_format_var.set("Custom")

style = ttk.Style()

tk.Label(root, text="Progress:").grid(row=3, column=0)
progress_bar = ttk.Progressbar(
    root, orient="horizontal", length=200, mode="determinate"
)
progress_bar.grid(row=3, columnspan=2, column=1)
tk.Button(root, text="Convert!", command=run_script).grid(row=2, column=2)

root.mainloop()
