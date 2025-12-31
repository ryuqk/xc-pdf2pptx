import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
import threading
import queue
import os
import sys
from dotenv import load_dotenv

# Import Core Logic
from pdf2pptx import DocumentProcessor, GeminiAnalyzer, PPTXBuilder

load_dotenv()

class PDF2PPTXApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF/Image to PPTX Converter (Gemini Powered)")
        self.geometry("800x600")
        
        self.file_queue = []
        self.processing = False
        self.cancel_event = threading.Event()
        self.msg_queue = queue.Queue()
        
        self._init_ui()
        self.after(100, self._process_queue)

    def _init_ui(self):
        # --- Top Frame: Settings ---
        settings_frame = ttk.LabelFrame(self, text="Settings", padding="10")
        settings_frame.pack(fill="x", padx=10, pady=5)
        
        # API Key
        ttk.Label(settings_frame, text="API Key:").pack(side="left")
        self.api_key_var = tk.StringVar(value=os.environ.get("GOOGLE_API_KEY", ""))
        self.api_key_entry = ttk.Entry(settings_frame, textvariable=self.api_key_var, width=40, show="*")
        self.api_key_entry.pack(side="left", padx=5)
        ttk.Button(settings_frame, text="Save", command=self._save_api_key, width=5).pack(side="left")
        
        # Mode
        ttk.Label(settings_frame, text="Mode:").pack(side="left", padx=10)
        self.mode_var = tk.StringVar(value="text_focus")
        modes = ["text_focus", "standard"]
        self.mode_combo = ttk.Combobox(settings_frame, textvariable=self.mode_var, values=modes, state="readonly", width=15)
        self.mode_combo.pack(side="left")

        # Font Scale
        self.font_scale_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(settings_frame, text="Font Scale (1.1x)", variable=self.font_scale_var).pack(side="left", padx=10)

        # --- Middle Frame: Drag & Drop List ---
        list_frame = ttk.LabelFrame(self, text="Files (Drag & Drop here)", padding="10")
        list_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.file_listbox = tk.Listbox(list_frame, selectmode="extended")
        self.file_listbox.pack(fill="both", expand=True, side="left")
        
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.file_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.file_listbox.config(yscrollcommand=scrollbar.set)
        
        # DND Checks
        self.file_listbox.drop_target_register(DND_FILES)
        self.file_listbox.dnd_bind('<<Drop>>', self._on_drop)
        
        # Buttons for list
        btn_frame = ttk.Frame(list_frame)
        btn_frame.pack(side="bottom", fill="x", pady=5)
        ttk.Button(btn_frame, text="Add Files...", command=self._add_files).pack(side="left")
        ttk.Button(btn_frame, text="Clear List", command=self._clear_list).pack(side="left", padx=5)
        
        # --- Output Settings & Run ---
        action_frame = ttk.Frame(self, padding="10")
        action_frame.pack(fill="x", padx=10)
        
        ttk.Label(action_frame, text="Output Folder:").pack(side="left")
        self.out_dir_var = tk.StringVar()
        ttk.Entry(action_frame, textvariable=self.out_dir_var, width=40).pack(side="left", padx=5)
        ttk.Button(action_frame, text="Browse...", command=self._browse_output).pack(side="left")
        
        # Prominent Start Button
        self.run_btn = tk.Button(action_frame, text="Start Conversion", command=self._start_processing, 
                                 bg="#0078D7", fg="white", font=("Meiryo UI", 11, "bold"), padx=20, pady=5)
        self.run_btn.pack(side="right", padx=5)
        
        # Cancel Button
        self.cancel_btn = tk.Button(action_frame, text="Cancel", command=self._cancel_processing,
                                    bg="#FFA500", fg="black", font=("Meiryo UI", 11, "bold"), padx=10, pady=5, state="disabled")
        self.cancel_btn.pack(side="right")
        
        # --- Progress & Log ---
        progress_frame = ttk.Frame(self, padding="10")
        progress_frame.pack(fill="x", padx=10)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill="x")
        
        log_frame = ttk.LabelFrame(self, text="Logs", padding="5")
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.log_area = scrolledtext.ScrolledText(log_frame, height=8, state="disabled")
        self.log_area.pack(fill="both", expand=True)

    def _on_drop(self, event):
        files = self.tk.splitlist(event.data)
        for f in files:
            if os.path.isfile(f):
                self.file_queue.append(f)
                self.file_listbox.insert(tk.END, f)

    def _add_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Documents", "*.pdf *.png *.jpg *.jpeg *.bmp")])
        for f in files:
            self.file_queue.append(f)
            self.file_listbox.insert(tk.END, f)

    def _clear_list(self):
        self.file_queue = []
        self.file_listbox.delete(0, tk.END)
        self.progress_var.set(0)

    def _save_api_key(self):
        key = self.api_key_var.get().strip()
        if not key:
            messagebox.showwarning("Warning", "API Key is empty.")
            return
            
        try:
            with open(".env", "w") as f:
                f.write(f"GOOGLE_API_KEY={key}\n")
            messagebox.showinfo("Success", "API Key saved to .env file.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save .env: {e}")

    def _browse_output(self):
        d = filedialog.askdirectory()
        if d:
            self.out_dir_var.set(d)

    def _log(self, msg):
        self.msg_queue.put(("log", msg))

    def _process_queue(self):
        while not self.msg_queue.empty():
            msg_type, content = self.msg_queue.get()
            if msg_type == "log":
                self.log_area.config(state="normal")
                self.log_area.insert(tk.END, content + "\n")
                self.log_area.see(tk.END)
                self.log_area.config(state="disabled")
            elif msg_type == "progress":
                self.progress_var.set(content)
            elif msg_type == "done":
                messagebox.showinfo("Complete", "All tasks completed successfully.")
                self.processing = False
                self.run_btn.config(state="normal")
                self.cancel_btn.config(state="disabled")
            elif msg_type == "cancelled":
                messagebox.showinfo("Cancelled", "Processing cancelled by user.")
                self.processing = False
                self.run_btn.config(state="normal")
                self.cancel_btn.config(state="disabled")
            elif msg_type == "error":
                messagebox.showerror("Error", content)
                self.processing = False
                self.run_btn.config(state="normal")
                self.cancel_btn.config(state="disabled")
        
        self.after(100, self._process_queue)

    def _cancel_processing(self):
        if self.processing:
            self.cancel_event.set()
            self.cancel_btn.config(state="disabled")
            self._log("Cancelling... please wait for current step to finish.")

    def _start_processing(self):
        if self.processing:
            return
        
        if not self.file_queue:
            messagebox.showwarning("Warning", "No files in the list.")
            return

        api_key = self.api_key_var.get().strip()
        if not api_key:
            messagebox.showerror("Error", "API Key is required.\nPlease enter a valid Gemini API Key in the Settings.")
            return
            
        self.processing = True
        self.cancel_event.clear()
        self.run_btn.config(state="disabled")
        self.cancel_btn.config(state="normal")
        self.progress_var.set(0)
        
        output_dir = self.out_dir_var.get()
        mode = self.mode_var.get()
        font_scale = 1.1 if self.font_scale_var.get() else 1.0
        
        thread = threading.Thread(target=self._worker, args=(self.file_queue, output_dir, mode, api_key, font_scale))
        thread.start()

    def _worker(self, files, output_dir, mode, api_key, font_scale):
        try:
            total_files = len(files)
            analyzer = GeminiAnalyzer(api_key)
            
            for i, file_path in enumerate(files):
                if self.cancel_event.is_set():
                    self._log("Conversion cancelled by user.")
                    break

                try:
                    self._log(f"Processing File {i+1}/{total_files}: {os.path.basename(file_path)}")
                    
                    # Determine output path
                    if output_dir:
                        base = os.path.basename(file_path)
                        name, _ = os.path.splitext(base)
                        out_path = os.path.join(output_dir, name + ".pptx")
                    else:
                        base = os.path.splitext(file_path)[0]
                        out_path = base + ".pptx"
                        
                    # Core Logic
                    proc = DocumentProcessor(file_path)
                    builder = PPTXBuilder(out_path, mode=mode, font_scale=font_scale)
                    
                    try:
                        num_pages = len(proc.doc)
                        for page_num in range(num_pages):
                            if self.cancel_event.is_set():
                                break

                            self._log(f"  - Page {page_num + 1}/{num_pages}...")
                            image, w, h = proc.get_page_image(page_num)
                            
                            if page_num == 0:
                                builder.set_slide_size(w/72, h/72)
                                try:
                                    builder.prs.slide_width = int(w * 12700)
                                    builder.prs.slide_height = int(h * 12700)
                                except:
                                    pass

                            layout_data = analyzer.analyze_page(image)
                            builder.add_slide(image, layout_data, w, h)
                            
                            # Progress
                            overall_progress = ((i) + (page_num+1)/num_pages) / total_files * 100
                            self.msg_queue.put(("progress", overall_progress))

                        if self.cancel_event.is_set():
                            self._log("Processing stopped for this file.")
                        else:
                            builder.save()
                            self._log(f"  - Saved to {out_path}")
                            
                    finally:
                        proc.close()
                    
                except Exception as e:
                    error_msg = str(e)
                    if "API key not valid" in error_msg or "API_KEY_INVALID" in error_msg:
                        self._log("Error: User provided an invalid API Key.")
                        self.msg_queue.put(("error", "Invalid API Key.\nPlease check your API Key in Settings and try again."))
                        return # Stop processing immediately
                    else:
                        self._log(f"Error processing {os.path.basename(file_path)}: {e}")
            
            if self.cancel_event.is_set():
                self.msg_queue.put(("progress", 0))
                self.msg_queue.put(("cancelled", ""))
            else:
                self.msg_queue.put(("progress", 100))
                self.msg_queue.put(("done", ""))

        except Exception as e:
            # ... (error handling code remains same)
            error_msg = str(e)
            if "API key not valid" in error_msg or "API_KEY_INVALID" in error_msg:
                 self.msg_queue.put(("error", "Invalid API Key.\nPlease check your API Key in Settings and try again."))
            else:
                 self.msg_queue.put(("error", f"Critical Error: {e}"))

if __name__ == "__main__":
    app = PDF2PPTXApp()
    app.mainloop()
