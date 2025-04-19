import os
import shutil
import zipfile
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.font import Font
import threading
import time

RAW_DIR = "raw"
TEMP_DIR = "temp"

def process_zst_files(zst_file_paths, raw_dir=RAW_DIR, temp_dir=TEMP_DIR, callback=None):
    if not os.path.exists(raw_dir):
        os.makedirs(raw_dir)
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)
    extracted_xls_files = []
    total_files = len([f for f in zst_file_paths if f.lower().endswith('.zst')])
    processed = 0

    for zst_path in zst_file_paths:
        filename = os.path.basename(zst_path)
        if not filename.lower().endswith(".zst"):
            continue

        if callback:
            callback(f"Processing {filename}...", processed / total_files)
        
        # Copy .ZST file to raw directory (if not already there)
        raw_zst_path = os.path.join(raw_dir, filename)
        if not os.path.exists(raw_zst_path):
            shutil.copy2(zst_path, raw_zst_path)

        # Work from the copy in raw
        zip_filename = filename[:-4] + ".zip"
        zip_path = os.path.join(raw_dir, zip_filename)

        # Copy and rename .zst to .zip in the raw directory
        shutil.copy2(raw_zst_path, zip_path)

        # Extract the .zip file
        extract_dir = os.path.join(temp_dir, filename[:-4] + "_extracted")
        try:
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)
        except Exception as e:
            if callback:
                callback(f"Error extracting {filename}: {e}", processed / total_files)
            continue

        # Find the 'data' file and rename it
        data_file_path = os.path.join(extract_dir, "data")
        if os.path.exists(data_file_path):
            new_xls_name = filename[:-4] + ".xls"
            new_xls_path = os.path.join(temp_dir, new_xls_name)
            shutil.move(data_file_path, new_xls_path)
            extracted_xls_files.append(new_xls_path)
        else:
            if callback:
                callback(f"No 'data' file found in {filename}", processed / total_files)

        # Clean up extracted folder and .zip file
        shutil.rmtree(extract_dir)
        os.remove(zip_path)
        
        processed += 1
        if callback:
            callback(f"Processed {processed} of {total_files} files", processed / total_files)

    return extracted_xls_files

def combine_xls_files(xls_file_paths, output_file="combined_output.xlsx", callback=None):
    combined_rows = []
    total_files = len(xls_file_paths)
    
    for i, file_path in enumerate(xls_file_paths):
        xls_file = os.path.basename(file_path)
        if callback:
            callback(f"Combining {xls_file}...", i / total_files)
            
        try:
            df = pd.read_excel(file_path, engine='xlrd')
        except Exception as e:
            if callback:
                callback(f"Error reading {xls_file}: {e}", i / total_files)
            continue

        filename_row = pd.DataFrame([[xls_file] + [''] * (df.shape[1] - 1)], columns=df.columns)
        combined_rows.append(filename_row)
        combined_rows.append(df)

    if combined_rows:
        if callback:
            callback("Creating final Excel file...", 0.9)
        combined_df = pd.concat(combined_rows, ignore_index=True)
        combined_df.to_excel(output_file, index=False)
        return True
    else:
        return False

class ModernButton(tk.Frame):
    def __init__(self, master=None, text="Button", command=None, width=None, state=tk.NORMAL, **kwargs):
        super().__init__(master, bg=master["bg"], **kwargs)
        
        self.default_bg = "#4a7abc"
        self.hover_bg = "#3a5d94"
        self.disabled_bg = "#cccccc"
        self.state = state
        self.command = command
        
        # Calculate width based on text if not specified
        if width is None:
            width = len(text) * 8 + 30
        
        # Create a canvas for the rounded rectangle
        self.canvas = tk.Canvas(
            self, 
            width=width, 
            height=36, 
            bg=master["bg"], 
            highlightthickness=0
        )
        self.canvas.pack()
        
        # Draw the rounded rectangle
        self.rect_id = self.canvas.create_rounded_rectangle(
            2, 2, width-2, 34, 
            radius=10, 
            fill=self.default_bg if state != tk.DISABLED else self.disabled_bg
        )
        
        # Add text
        self.text_id = self.canvas.create_text(
            width//2, 18, 
            text=text, 
            fill="white", 
            font=("Segoe UI", 10, "bold")
        )
        
        # Bind events
        self.canvas.bind("<Enter>", self._on_enter)
        self.canvas.bind("<Leave>", self._on_leave)
        self.canvas.bind("<Button-1>", self._on_click)
        
    def _on_enter(self, event):
        if self.state != tk.DISABLED:
            self.canvas.itemconfig(self.rect_id, fill=self.hover_bg)
            self.canvas.config(cursor="hand2")
    
    def _on_leave(self, event):
        if self.state != tk.DISABLED:
            self.canvas.itemconfig(self.rect_id, fill=self.default_bg)
            self.canvas.config(cursor="")
    
    def _on_click(self, event):
        if self.state != tk.DISABLED and self.command:
            self.command()
    
    def config(self, **kwargs):
        if "state" in kwargs:
            self.state = kwargs["state"]
            if self.state == tk.DISABLED:
                self.canvas.itemconfig(self.rect_id, fill=self.disabled_bg)
            else:
                self.canvas.itemconfig(self.rect_id, fill=self.default_bg)
        if "command" in kwargs:
            self.command = kwargs["command"]
        if "text" in kwargs:
            self.canvas.itemconfig(self.text_id, text=kwargs["text"])

# Add rounded rectangle method to Canvas
tk.Canvas.create_rounded_rectangle = lambda self, x1, y1, x2, y2, radius=25, **kwargs: self.create_polygon(
    x1+radius, y1,
    x2-radius, y1,
    x2, y1,
    x2, y1+radius,
    x2, y2-radius,
    x2, y2,
    x2-radius, y2,
    x1+radius, y2,
    x1, y2,
    x1, y2-radius,
    x1, y1+radius,
    x1, y1,
    smooth=True, **kwargs
)

class FileListFrame(tk.Frame):
    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)
        self.config(bg='white', highlightbackground="#e0e0e0", highlightthickness=1)
        
        # Create a canvas with scrollbar
        self.canvas = tk.Canvas(self, bg='white', highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg='white')
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas_frame = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Pack the canvas and scrollbar
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # Bind canvas resize to window resize
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        
        # Empty label to show when no files
        self.empty_label = tk.Label(self.scrollable_frame, text="No files selected", bg='white', fg='#999')
        self.empty_label.pack(pady=20)
        
        self.file_labels = []
    
    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_frame, width=event.width)
    
    def set_files(self, file_paths):
        # Clear existing labels
        for label in self.file_labels:
            label.destroy()
        self.file_labels = []
        
        if not file_paths:
            self.empty_label.pack(pady=20)
            return
        
        self.empty_label.pack_forget()
        
        # Add new file labels
        for i, file_path in enumerate(file_paths):
            filename = os.path.basename(file_path)
            bg_color = '#f5f5f5' if i % 2 == 0 else 'white'
            
            frame = tk.Frame(self.scrollable_frame, bg=bg_color)
            frame.pack(fill='x', expand=True)
            
            label = tk.Label(frame, text=filename, anchor='w', bg=bg_color, padx=10, pady=5)
            label.pack(side='left', fill='x', expand=True)
            
            self.file_labels.append(frame)

class ZSTCombineApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ZST to Excel Combiner")
        self.root.configure(bg='#f0f0f0')
        self.root.resizable(True, True)
        self.center_window(600, 500)
        
        # Set application style
        self.style = ttk.Style()
        self.style.configure('TProgressbar', thickness=15, troughcolor='#f0f0f0', background='#4a7abc')
        
        # Main frame
        main_frame = tk.Frame(root, bg='#f0f0f0')
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Title label
        title_font = Font(family="Segoe UI", size=18, weight="bold")
        title_label = tk.Label(main_frame, text="ZST to Excel Combiner", font=title_font, bg='#f0f0f0')
        title_label.pack(pady=(0, 20))
        
        # Button frame - centered
        button_frame = tk.Frame(main_frame, bg='#f0f0f0')
        button_frame.pack(fill='x', pady=(0, 20))
        
        # Center container for buttons
        center_frame = tk.Frame(button_frame, bg='#f0f0f0')
        center_frame.pack(anchor='center')
        
        self.select_button = ModernButton(center_frame, text="Select Files", command=self.select_files)
        self.select_button.pack(side='left', padx=10)
        
        self.combine_button = ModernButton(center_frame, text="Extract, Combine and Save", 
                                          command=self.start_processing, state=tk.DISABLED)
        self.combine_button.pack(side='left', padx=10)
        
        # File count label
        self.file_count_label = tk.Label(main_frame, text="No files selected", bg='#f0f0f0', anchor='w')
        self.file_count_label.pack(fill='x', pady=(0, 10))
        
        # File list frame
        list_frame_label = tk.Label(main_frame, text="Selected Files:", bg='#f0f0f0', anchor='w')
        list_frame_label.pack(fill='x')
        
        self.file_list_frame = FileListFrame(main_frame)
        self.file_list_frame.pack(fill='both', expand=True, pady=(5, 15))
        
        # Status frame
        status_frame = tk.Frame(main_frame, bg='#f0f0f0')
        status_frame.pack(fill='x', pady=(10, 0))
        
        self.status_label = tk.Label(status_frame, text="Ready", bg='#f0f0f0', anchor='w')
        self.status_label.pack(fill='x')
        
        self.progress_bar = ttk.Progressbar(status_frame, orient='horizontal', mode='determinate')
        self.progress_bar.pack(fill='x', pady=(5, 0))
        
        self.selected_zst_files = []
        self.processing_thread = None

    def center_window(self, width, height):
        self.root.update_idletasks()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def select_files(self):
        file_paths = filedialog.askopenfilenames(
            title="Select Files",
            filetypes=[("All files", "*.*"), ("ZST files", "*.zst")]
        )
        if file_paths:
            self.selected_zst_files = list(file_paths)
            self.file_count_label.config(text=f"{len(self.selected_zst_files)} files selected")
            self.file_list_frame.set_files(self.selected_zst_files)
            self.combine_button.config(state=tk.NORMAL)
            self.status_label.config(text="Ready to process")
            self.progress_bar['value'] = 0
        else:
            self.selected_zst_files = []
            self.file_count_label.config(text="No files selected")
            self.file_list_frame.set_files([])
            self.combine_button.config(state=tk.DISABLED)
            self.status_label.config(text="Ready")

    def update_progress(self, message, progress):
        def update():
            self.status_label.config(text=message)
            self.progress_bar['value'] = progress * 100
        self.root.after(0, update)

    def start_processing(self):
        if not self.selected_zst_files:
            messagebox.showwarning("No Files", "Please select files first.")
            return
        
        # Disable buttons during processing
        self.select_button.config(state=tk.DISABLED)
        self.combine_button.config(state=tk.DISABLED)
        
        # Ask for output file location
        output_file = filedialog.asksaveasfilename(
            title="Save Combined Excel File",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not output_file:
            self.select_button.config(state=tk.NORMAL)
            self.combine_button.config(state=tk.NORMAL)
            self.status_label.config(text="Save cancelled")
            return
        
        # Start processing in a separate thread
        self.processing_thread = threading.Thread(
            target=self.process_files,
            args=(output_file,)
        )
        self.processing_thread.daemon = True
        self.processing_thread.start()

    def process_files(self, output_file):
        try:
            # Clean temp dir before extracting
            self.update_progress("Preparing temporary directory...", 0.05)
            if os.path.exists(TEMP_DIR):
                shutil.rmtree(TEMP_DIR)
            os.makedirs(TEMP_DIR, exist_ok=True)
            
            # Extract files
            self.update_progress("Extracting files...", 0.1)
            xls_files = process_zst_files(
                self.selected_zst_files, 
                RAW_DIR, 
                TEMP_DIR, 
                callback=self.update_progress
            )
            
            if not xls_files:
                self.root.after(0, lambda: messagebox.showerror(
                    "Error", 
                    "No .xls files were extracted from the selected files."
                ))
                self.root.after(0, self.reset_ui)
                return
            
            # Combine files
            self.update_progress("Combining files...", 0.7)
            success = combine_xls_files(
                xls_files, 
                output_file, 
                callback=self.update_progress
            )
            
            if success:
                # Clean up temp folder
                self.update_progress("Cleaning up temporary files...", 0.95)
                try:
                    shutil.rmtree(TEMP_DIR)
                except Exception as e:
                    print(f"Error deleting temp folder: {e}")
                
                self.update_progress(f"Completed! File saved as {os.path.basename(output_file)}", 1.0)
                self.root.after(0, lambda: messagebox.showinfo(
                    "Success", 
                    f"Combined file saved as:\n{output_file}"
                ))
            else:
                self.root.after(0, lambda: messagebox.showerror(
                    "Error", 
                    "No files were combined."
                ))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror(
                "Error", 
                f"An error occurred: {str(e)}"
            ))
        finally:
            self.root.after(0, self.reset_ui)
    
    def reset_ui(self):
        self.select_button.config(state=tk.NORMAL)
        if self.selected_zst_files:
            self.combine_button.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = ZSTCombineApp(root)
    root.mainloop()
