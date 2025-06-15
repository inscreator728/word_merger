import os
from tkinter import Tk, filedialog, messagebox, ttk, IntVar, StringVar
from tkinter.scrolledtext import ScrolledText
from docx import Document
import time
import threading

class MergeWordApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Word Document Merger")
        self.root.geometry("700x500")
        self.root.resizable(False, False)

        self.file_paths = []
        self.order_var = StringVar()

        # GUI Components
        self.setup_gui()

    def setup_gui(self):
        ttk.Label(self.root, text="Word Document Merger", font=("Arial", 16)).pack(pady=10)

        # File List Display
        self.file_list = ScrolledText(self.root, width=80, height=8, state='disabled', wrap='word')
        self.file_list.pack(pady=10)

        # Buttons
        button_frame = ttk.Frame(self.root)
        button_frame.pack(pady=10)

        ttk.Button(button_frame, text="Add Files", command=self.browse_multiple_files).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="Clear List", command=self.clear_list).grid(row=0, column=1, padx=5)
        ttk.Button(button_frame, text="Merge Files", command=self.start_merge).grid(row=0, column=2, padx=5)

        # Order Entry
        ttk.Label(self.root, text="Edit Order (Comma-separated indices):", font=("Arial", 10)).pack(pady=5)
        self.order_entry = ttk.Entry(self.root, textvariable=self.order_var, width=50)
        self.order_entry.pack(pady=5)

        # Progress Bar
        self.progress_var = IntVar()
        self.progress_bar = ttk.Progressbar(self.root, orient="horizontal", length=500, mode="determinate", variable=self.progress_var)
        self.progress_bar.pack(pady=10)

        # Status and History
        self.status_label = ttk.Label(self.root, text="", font=("Arial", 10), foreground="green")
        self.status_label.pack(pady=5)

        ttk.Label(self.root, text="Merge History:", font=("Arial", 12)).pack(pady=5)
        self.history = ScrolledText(self.root, width=80, height=8, state='disabled', wrap='word')
        self.history.pack(pady=10)

    def browse_multiple_files(self):
        files = filedialog.askopenfilenames(
            title="Select Word Documents",
            filetypes=[("Word Documents", "*.docx")]
        )
        if files:
            self.file_paths.extend(files)
            self.update_file_list()

    def clear_list(self):
        self.file_paths = []
        self.update_file_list()
        self.order_var.set("")

    def update_file_list(self):
        self.file_list.config(state='normal')
        self.file_list.delete(1.0, 'end')
        for index, file_path in enumerate(self.file_paths, start=1):
            self.file_list.insert('end', f"{index}. {file_path}\n")
        self.file_list.config(state='disabled')
        self.order_var.set(",".join(map(str, range(1, len(self.file_paths) + 1))))

    def start_merge(self):
        if not self.file_paths:
            messagebox.showwarning("No Files", "No files selected to merge.")
            return

        try:
            order = list(map(int, self.order_var.get().split(",")))
            ordered_files = [self.file_paths[i - 1] for i in order]
        except (ValueError, IndexError):
            messagebox.showerror("Error", "Invalid order input.")
            return

        output_path = filedialog.asksaveasfilename(
            title="Save Merged Document",
            defaultextension=".docx",
            filetypes=[("Word Documents", "*.docx")]
        )

        if not output_path:
            return

        threading.Thread(target=self.merge_files, args=(output_path, ordered_files)).start()

    def merge_files(self, output_path, ordered_files):
        try:
            total_files = len(ordered_files)
            self.progress_var.set(0)
            self.progress_bar["maximum"] = total_files

            merged_document = Document()
            start_time = time.time()

            for i, file_path in enumerate(ordered_files):
                doc = Document(file_path)
                for element in doc.element.body:
                    merged_document.element.body.append(element)

                self.progress_var.set(i + 1)
                elapsed_time = time.time() - start_time
                estimated_total_time = (elapsed_time / (i + 1)) * total_files
                remaining_time = estimated_total_time - elapsed_time
                self.status_label.config(
                    text=f"Merging {i + 1}/{total_files} files... {self.progress_var.get()}% complete. ETA: {remaining_time:.1f} seconds"
                )
                self.update_history(f"Merged: {file_path}")
                time.sleep(0.5)  # Simulating processing time for better progress visibility

            merged_document.save(output_path)
            self.status_label.config(text="Merge Completed Successfully!", foreground="green")
            messagebox.showinfo("Success", f"Documents merged successfully!\nSaved at: {output_path}")
        except Exception as e:
            self.status_label.config(text="Error during merge. Check log.", foreground="red")
            messagebox.showerror("Error", f"Error merging documents: {e}")

    def update_history(self, message):
        self.history.config(state='normal')
        self.history.insert('end', f"{message}\n")
        self.history.config(state='disabled')

if __name__ == "__main__":
    root = Tk()
    app = MergeWordApp(root)
    root.mainloop()
