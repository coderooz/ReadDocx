import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document
import pyttsx3
import threading
import time

class DocxReader:
    def __init__(self, master):
        self.master = master
        master.title("DOCX Reader")

        # Initialize variables
        self.file_path = ""
        self.text_chunks = []
        self.engine = pyttsx3.init()
        self.is_paused = False
        self.is_stopped = False
        self.current_index = 0

        # Create UI elements
        self.load_button = tk.Button(master, text="Load DOCX", command=self.load_file)
        self.load_button.pack(pady=5)

        self.start_label = tk.Label(master, text="Start Page (paragraph index):")
        self.start_label.pack()
        self.start_entry = tk.Entry(master)
        self.start_entry.pack(pady=2)

        self.end_label = tk.Label(master, text="End Page (paragraph index):")
        self.end_label.pack()
        self.end_entry = tk.Entry(master)
        self.end_entry.pack(pady=2)

        self.read_button = tk.Button(master, text="Read", command=self.start_reading)
        self.read_button.pack(pady=5)

        # Control buttons
        self.pause_button = tk.Button(master, text="Pause", command=self.pause)
        self.pause_button.pack(side=tk.LEFT, padx=10, pady=10)

        self.resume_button = tk.Button(master, text="Resume", command=self.resume)
        self.resume_button.pack(side=tk.LEFT, padx=10, pady=10)

        self.stop_button = tk.Button(master, text="Stop", command=self.stop)
        self.stop_button.pack(side=tk.LEFT, padx=10, pady=10)

    def load_file(self):
        # Open file dialog to choose a DOCX file
        self.file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if self.file_path:
            doc = Document(self.file_path)
            # Treat each non-empty paragraph as a "page"
            self.text_chunks = [para.text for para in doc.paragraphs if para.text.strip() != ""]
            messagebox.showinfo("Loaded", f"Loaded {len(self.text_chunks)} pages (non-empty paragraphs).")

    def start_reading(self):
        try:
            start_page = int(self.start_entry.get())
            end_page = int(self.end_entry.get())
        except ValueError:
            messagebox.showerror("Error", "Please enter valid numeric page numbers.")
            return

        # Validate page range
        if start_page < 1 or end_page > len(self.text_chunks) or start_page > end_page:
            messagebox.showerror("Error", "Invalid page range.")
            return

        # Reset control flags and set starting index (convert to zero-based)
        self.is_paused = False
        self.is_stopped = False
        self.current_index = start_page - 1

        # Start the reading process in a separate thread
        threading.Thread(target=self.read_text, args=(end_page,), daemon=True).start()

    def read_text(self, end_page):
        while self.current_index < end_page and not self.is_stopped:
            # Wait if paused
            while self.is_paused and not self.is_stopped:
                time.sleep(0.1)
            if self.is_stopped:
                break
            # Read current "page"
            text = self.text_chunks[self.current_index]
            self.engine.say(text)
            self.engine.runAndWait()
            self.current_index += 1

    def pause(self):
        self.is_paused = True

    def resume(self):
        self.is_paused = False

    def stop(self):
        self.is_stopped = True
        self.engine.stop()

if __name__ == "__main__":
    root = tk.Tk()
    app = DocxReader(root)
    root.mainloop()
