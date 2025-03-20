import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image
import PyPDF2
import fitz  # PyMuPDF
import tempfile
from PIL import Image, ImageTk  # Add ImageTk here

class PDFToolsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Toolkit")
        self.root.geometry("900x700")  # Larger initial size
        self.root.minsize(600, 500)    # Set minimum window size
        
        # Configure row and column weights to make UI expandable
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=0)
        
        # Set theme for a modern look
        style = ttk.Style()
        style.theme_use('clam')  # Use a more modern theme if available
        
        # Configure colors for better visual appearance
        bg_color = "#f5f5f5"
        accent_color = "#3498db"
        
        style.configure('TButton', font=('Arial', 10), padding=5)
        style.configure('TFrame', background=bg_color)
        style.configure('TLabel', background=bg_color, font=('Arial', 10))
        style.configure('TNotebook', background=bg_color, tabposition='n')
        style.map('TButton', background=[('active', accent_color)])
        
        # Create notebook (tabs) with grid instead of pack for better responsiveness
        self.notebook = ttk.Notebook(root)
        self.notebook.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        # Create tabs
        self.create_image_to_pdf_tab()
        self.create_pdf_merger_tab()
        self.create_pdf_splitter_tab()
        self.create_pdf_page_remover_tab()
        self.create_pdf_fill_sign_tab()
        
        # Status bar with grid
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        self.status_bar = tk.Label(root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.grid(row=1, column=0, sticky="ew")
        
    def create_image_to_pdf_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Image to PDF")
        
        # Configure the grid system for responsiveness
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(0, weight=1)
        tab.rowconfigure(1, weight=0)
        
        # Frame for listbox with a scrollbar
        list_frame = ttk.Frame(tab)
        list_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # Variables
        self.image_files = []
        
        # Create scrollbar for listbox
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        self.image_listbox = tk.Listbox(list_frame, width=70, height=15, yscrollcommand=scrollbar.set)
        self.image_listbox.grid(row=0, column=0, sticky="nsew")
        scrollbar.config(command=self.image_listbox.yview)
        
        # Buttons frame
        btn_frame = ttk.Frame(tab)
        btn_frame.grid(row=1, column=0, pady=10, sticky="ew")
        
        # Center buttons by configuring columns
        for i in range(4):
            btn_frame.columnconfigure(i, weight=1)
        
        ttk.Button(btn_frame, text="Add Images", command=self.add_images).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text="Remove Selected", command=self.remove_selected_image).grid(row=0, column=1, padx=5)
        ttk.Button(btn_frame, text="Clear All", command=self.clear_images).grid(row=0, column=2, padx=5)
        ttk.Button(btn_frame, text="Convert to PDF", command=self.convert_images_to_pdf).grid(row=0, column=3, padx=5)
        
    def create_pdf_merger_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="PDF Merger")
        
        # Configure the grid system for responsiveness
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(0, weight=1)
        tab.rowconfigure(1, weight=0)
        
        # Frame for listbox with a scrollbar
        list_frame = ttk.Frame(tab)
        list_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # Variables
        self.pdf_files = []
        
        # Create scrollbar for listbox
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        self.pdf_listbox = tk.Listbox(list_frame, width=70, height=15, yscrollcommand=scrollbar.set)
        self.pdf_listbox.grid(row=0, column=0, sticky="nsew")
        scrollbar.config(command=self.pdf_listbox.yview)
        
        # Buttons frame
        btn_frame = ttk.Frame(tab)
        btn_frame.grid(row=1, column=0, pady=10, sticky="ew")
        
        # Arrange buttons in two rows for better layout on smaller screens
        btn_frame.columnconfigure((0, 1, 2), weight=1)
        
        ttk.Button(btn_frame, text="Add PDFs", command=self.add_pdfs).grid(row=0, column=0, padx=5, pady=3)
        ttk.Button(btn_frame, text="Remove Selected", command=self.remove_selected_pdf).grid(row=0, column=1, padx=5, pady=3)
        ttk.Button(btn_frame, text="Clear All", command=self.clear_pdfs).grid(row=0, column=2, padx=5, pady=3)
        
        ttk.Button(btn_frame, text="Move Up", command=self.move_pdf_up).grid(row=1, column=0, padx=5, pady=3)
        ttk.Button(btn_frame, text="Move Down", command=self.move_pdf_down).grid(row=1, column=1, padx=5, pady=3)
        ttk.Button(btn_frame, text="Merge PDFs", command=self.merge_pdfs).grid(row=1, column=2, padx=5, pady=3)
        
    def create_pdf_splitter_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="PDF Splitter")
        
        # Frame for file selection
        file_frame = ttk.Frame(tab)
        file_frame.pack(pady=10, fill=tk.X, padx=10)
        
        ttk.Label(file_frame, text="PDF File:").grid(row=0, column=0, padx=5, pady=5)
        self.split_pdf_path = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.split_pdf_path, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_pdf_to_split).grid(row=0, column=2, padx=5, pady=5)
        
        # Frame for splitting options
        options_frame = ttk.LabelFrame(tab, text="Splitting Options")
        options_frame.pack(pady=10, padx=10, fill=tk.X)
        
        # Split by range
        self.split_option = tk.StringVar(value="range")
        ttk.Radiobutton(options_frame, text="Split by Range", variable=self.split_option, value="range").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        
        range_frame = ttk.Frame(options_frame)
        range_frame.grid(row=1, column=0, columnspan=2, padx=20, pady=5, sticky=tk.W)
        
        ttk.Label(range_frame, text="Page Range (e.g., 1-3,5,7-9):").pack(side=tk.LEFT, padx=5)
        self.page_range = tk.StringVar()
        ttk.Entry(range_frame, textvariable=self.page_range, width=30).pack(side=tk.LEFT, padx=5)
        
        # Split each page
        ttk.Radiobutton(options_frame, text="Split each page into separate PDF", variable=self.split_option, value="each").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        
        # Button to execute
        ttk.Button(tab, text="Split PDF", command=self.split_pdf).pack(pady=20)
        
    def create_pdf_page_remover_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Page Remover")
        
        # Frame for file selection
        file_frame = ttk.Frame(tab)
        file_frame.pack(pady=10, fill=tk.X, padx=10)
        
        ttk.Label(file_frame, text="PDF File:").grid(row=0, column=0, padx=5, pady=5)
        self.remove_pdf_path = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.remove_pdf_path, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_pdf_to_remove_pages).grid(row=0, column=2, padx=5, pady=5)
        
        # Frame for page info
        info_frame = ttk.Frame(tab)
        info_frame.pack(pady=10, fill=tk.X, padx=10)
        
        self.total_pages_var = tk.StringVar(value="Total Pages: 0")
        ttk.Label(info_frame, textvariable=self.total_pages_var).pack(pady=5)
        
        # Frame for page selection
        page_frame = ttk.LabelFrame(tab, text="Pages to Remove")
        page_frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        
        ttk.Label(page_frame, text="Enter page numbers to remove (e.g., 1-3,5,7-9):").pack(anchor=tk.W, padx=5, pady=5)
        self.pages_to_remove = tk.StringVar()
        ttk.Entry(page_frame, textvariable=self.pages_to_remove, width=50).pack(padx=5, pady=5, fill=tk.X)
        
        # Button to execute
        ttk.Button(tab, text="Remove Pages", command=self.remove_pdf_pages).pack(pady=20)
        
    def create_pdf_fill_sign_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Fill & Sign")
        
        # Make the tab responsive
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(2, weight=1)  # Make the canvas area expand
        
        # Frame for file selection
        file_frame = ttk.Frame(tab)
        file_frame.grid(row=0, column=0, pady=10, padx=10, sticky="ew")
        file_frame.columnconfigure(1, weight=1)
        
        ttk.Label(file_frame, text="PDF File:").grid(row=0, column=0, padx=5, pady=5)
        self.sign_pdf_path = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.sign_pdf_path).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(file_frame, text="Browse", command=self.browse_pdf_to_sign).grid(row=0, column=2, padx=5, pady=5)

        # Frame for signature
        sign_frame = ttk.LabelFrame(tab, text="Signature")
        sign_frame.grid(row=1, column=0, pady=5, padx=10, sticky="ew")
        sign_frame.columnconfigure(0, weight=1)
        
        ttk.Button(sign_frame, text="Add Signature Image", command=self.add_signature_image).grid(row=0, column=0, pady=5)
        self.signature_path = tk.StringVar()
        ttk.Label(sign_frame, textvariable=self.signature_path).grid(row=1, column=0, pady=5)

        # Canvas area in a frame that expands with window
        sig_place_frame = ttk.LabelFrame(tab, text="PDF Preview")
        sig_place_frame.grid(row=2, column=0, pady=5, padx=10, sticky="nsew")
        sig_place_frame.columnconfigure(0, weight=1)
        sig_place_frame.rowconfigure(0, weight=1)
        
        # Canvas that resizes with window
        self.pdf_canvas = tk.Canvas(sig_place_frame, bg='white')
        self.pdf_canvas.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        
        # Add controls for page navigation
        nav_frame = ttk.Frame(sig_place_frame)
        nav_frame.grid(row=1, column=0, sticky="ew")
        nav_frame.columnconfigure((0, 1, 2), weight=1)
        
        ttk.Button(nav_frame, text="Previous Page", command=self.show_previous_page).grid(row=0, column=0, padx=5, pady=5)
        self.page_label = ttk.Label(nav_frame, text="Page: 1")
        self.page_label.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(nav_frame, text="Next Page", command=self.show_next_page).grid(row=0, column=2, padx=5, pady=5)
        
        # Place signature button
        ttk.Button(sig_place_frame, text="Place Signature", command=self.enter_signature_mode).grid(row=2, column=0, pady=5)
        
        # Button to save the final PDF
        ttk.Button(tab, text="Save PDF with Fields and Signature", command=self.save_filled_pdf).grid(row=3, column=0, pady=10)
        
        # Bind window resize event to update canvas
        self.pdf_canvas.bind("<Configure>", self.on_canvas_resize)

    def load_pdf_preview(self):
        pdf_path = self.sign_pdf_path.get()
        if not pdf_path:
            messagebox.showwarning("Warning", "Please select a PDF file first")
            return False
        
        try:
            self.doc = fitz.open(pdf_path)
            self.current_page = 0
            self.render_pdf_page()
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Error loading PDF: {str(e)}")
            return False

    def render_pdf_page(self):
        if not hasattr(self, 'doc'):
            return
    
    # Clear the canvas
        self.pdf_canvas.delete("all")
        
        # Get the current page
        page = self.doc[self.current_page]
        
        # Render the page to an image
        pix = page.get_pixmap(matrix=fitz.Matrix(1, 1))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        # Get current canvas dimensions
        self.pdf_canvas.update()  # Ensure canvas is updated
        canvas_width = self.pdf_canvas.winfo_width() or 500  # Default if not yet rendered
        canvas_height = self.pdf_canvas.winfo_height() or 700
        
        # Adjust image size to fit canvas while maintaining aspect ratio
        img_width, img_height = img.size
        scale = min(canvas_width/img_width, canvas_height/img_height) * 0.95  # 5% margin
        new_width = int(img_width * scale)
        new_height = int(img_height * scale)
        img = img.resize((new_width, new_height), Image.LANCZOS)
        
        # Convert to PhotoImage and keep a reference
        self.pdf_image = ImageTk.PhotoImage(img)
        self.pdf_canvas.create_image(canvas_width//2, canvas_height//2, image=self.pdf_image, anchor=tk.CENTER)
        
        # Update page label
        self.page_label.config(text=f"Page: {self.current_page + 1} of {len(self.doc)}")
        
        # Store the image dimensions for coordinate conversion
        self.img_width = img_width
        self.img_height = img_height
        self.display_width = new_width
        self.display_height = new_height

    def show_previous_page(self):
        if hasattr(self, 'doc') and self.current_page > 0:
            self.current_page -= 1
            self.render_pdf_page()

    def show_next_page(self):
        if hasattr(self, 'doc') and self.current_page < len(self.doc) - 1:
            self.current_page += 1
            self.render_pdf_page()

    def enter_signature_mode(self):
        if not hasattr(self, '_signature_file'):
            messagebox.showwarning("Warning", "Please add a signature image first")
            return
        
        if not hasattr(self, 'doc'):
            if not self.load_pdf_preview():
                return
    
    # Change cursor to indicate signature placement mode
        self.pdf_canvas.config(cursor="crosshair")
    
    # Load the signature image as a preview
        self.signature_img = Image.open(self._signature_file)
        self.signature_img = self.signature_img.resize((200, int(200 * self.signature_img.height / self.signature_img.width)), Image.LANCZOS)
        self.signature_tk = ImageTk.PhotoImage(self.signature_img)
    
    # Bind mouse events for signature placement
        self.pdf_canvas.bind("<Motion>", self.move_signature)
        self.pdf_canvas.bind("<Button-1>", self.place_signature_on_click)
    
        self.status_var.set("Click to place signature")

    def move_signature(self, event):
    # Display signature preview at mouse position
        if hasattr(self, 'signature_preview_id'):
            self.pdf_canvas.delete(self.signature_preview_id)
    
        self.signature_preview_id = self.pdf_canvas.create_image(
            event.x, event.y, 
            image=self.signature_tk, 
            anchor=tk.CENTER,
            tags="signature_preview"
    )

    def place_signature_on_click(self, event):
        if not hasattr(self, 'doc'):
            return
        
    # Convert canvas coordinates to PDF coordinates
        canvas_width = self.pdf_canvas.winfo_width()
        canvas_height = self.pdf_canvas.winfo_height()
    
    # Calculate offsets due to centering
        offset_x = (canvas_width - self.display_width) / 2
        offset_y = (canvas_height - self.display_height) / 2
    
    # Adjust click position by offsets
        pdf_x = ((event.x - offset_x) / self.display_width) * self.img_width
        pdf_y = ((event.y - offset_y) / self.display_height) * self.img_height
    
    # Add the signature to the PDF
        page = self.doc[self.current_page]
        signature_width = 200  # Default width
        aspect_ratio = self.signature_img.height / self.signature_img.width
        signature_height = signature_width * aspect_ratio
    
    # Center the signature at the click position
        rect = fitz.Rect(
            pdf_x - signature_width/2, 
            pdf_y - signature_height/2, 
            pdf_x + signature_width/2, 
            pdf_y + signature_height/2
        )
    
        page.insert_image(rect, filename=self._signature_file)
    
    # Update the preview
        self.render_pdf_page()
    
    # Reset cursor and unbind events
        self.pdf_canvas.config(cursor="")
        self.pdf_canvas.unbind("<Motion>")
        self.pdf_canvas.unbind("<Button-1>")
    
        if hasattr(self, 'signature_preview_id'):
            self.pdf_canvas.delete(self.signature_preview_id)
    
        self.status_var.set("Signature placed on PDF")

    # Image to PDF methods
    def add_images(self):
        files = filedialog.askopenfilenames(
            title="Select Images",
            filetypes=[("Image files", "*.jpg *.jpeg *.png *.bmp *.tiff")]
        )
        
        if files:
            for file in files:
                if file not in self.image_files:
                    self.image_files.append(file)
                    self.image_listbox.insert(tk.END, os.path.basename(file))
            
            self.status_var.set(f"{len(files)} images added")
    
    def remove_selected_image(self):
        try:
            selected_idx = self.image_listbox.curselection()[0]
            self.image_listbox.delete(selected_idx)
            self.image_files.pop(selected_idx)
            self.status_var.set("Image removed")
        except IndexError:
            messagebox.showwarning("Warning", "No image selected")
    
    def clear_images(self):
        self.image_listbox.delete(0, tk.END)
        self.image_files.clear()
        self.status_var.set("All images cleared")
    
    def convert_images_to_pdf(self):
        if not self.image_files:
            messagebox.showwarning("Warning", "No images added")
            return
            
        output_file = filedialog.asksaveasfilename(
            title="Save PDF As",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")]
        )
        
        if output_file:
            try:
                # Open first image to get size
                first_image = Image.open(self.image_files[0])
                img_width, img_height = first_image.size
                
                # Create a PDF with the same dimensions as the first image
                images = []
                for img_path in self.image_files:
                    img = Image.open(img_path)
                    if img.mode == 'RGBA':
                        img = img.convert('RGB')
                    images.append(img)
                
                # Save all images as pages in the PDF
                images[0].save(
                    output_file,
                    save_all=True,
                    append_images=images[1:],
                    resolution=100.0
                )
                
                self.status_var.set(f"PDF created successfully: {os.path.basename(output_file)}")
                messagebox.showinfo("Success", f"PDF created successfully:\n{output_file}")
            except Exception as e:
                messagebox.showerror("Error", f"Error creating PDF: {str(e)}")
    
    # PDF Merger methods
    def add_pdfs(self):
        files = filedialog.askopenfilenames(
            title="Select PDFs",
            filetypes=[("PDF files", "*.pdf")]
        )
        
        if files:
            for file in files:
                if file not in self.pdf_files:
                    self.pdf_files.append(file)
                    self.pdf_listbox.insert(tk.END, os.path.basename(file))
            
            self.status_var.set(f"{len(files)} PDFs added")
    
    def remove_selected_pdf(self):
        try:
            selected_idx = self.pdf_listbox.curselection()[0]
            self.pdf_listbox.delete(selected_idx)
            self.pdf_files.pop(selected_idx)
            self.status_var.set("PDF removed")
        except IndexError:
            messagebox.showwarning("Warning", "No PDF selected")
    
    def clear_pdfs(self):
        self.pdf_listbox.delete(0, tk.END)
        self.pdf_files.clear()
        self.status_var.set("All PDFs cleared")
    
    def move_pdf_up(self):
        try:
            selected_idx = self.pdf_listbox.curselection()[0]
            if selected_idx > 0:
                # Swap in list
                self.pdf_files[selected_idx], self.pdf_files[selected_idx-1] = self.pdf_files[selected_idx-1], self.pdf_files[selected_idx]
                
                # Update listbox
                text = self.pdf_listbox.get(selected_idx)
                self.pdf_listbox.delete(selected_idx)
                self.pdf_listbox.insert(selected_idx-1, text)
                self.pdf_listbox.selection_set(selected_idx-1)
        except IndexError:
            messagebox.showwarning("Warning", "No PDF selected")
    
    def move_pdf_down(self):
        try:
            selected_idx = self.pdf_listbox.curselection()[0]
            if selected_idx < len(self.pdf_files) - 1:
                # Swap in list
                self.pdf_files[selected_idx], self.pdf_files[selected_idx+1] = self.pdf_files[selected_idx+1], self.pdf_files[selected_idx]
                
                # Update listbox
                text = self.pdf_listbox.get(selected_idx)
                self.pdf_listbox.delete(selected_idx)
                self.pdf_listbox.insert(selected_idx+1, text)
                self.pdf_listbox.selection_set(selected_idx+1)
        except IndexError:
            messagebox.showwarning("Warning", "No PDF selected")
    
    def merge_pdfs(self):
        if len(self.pdf_files) < 2:
            messagebox.showwarning("Warning", "Add at least 2 PDFs to merge")
            return
            
        output_file = filedialog.asksaveasfilename(
            title="Save Merged PDF As",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")]
        )
        
        if output_file:
            try:
                pdf_merger = PyPDF2.PdfMerger()
                
                for pdf_file in self.pdf_files:
                    pdf_merger.append(pdf_file)
                
                with open(output_file, 'wb') as f:
                    pdf_merger.write(f)
                
                self.status_var.set(f"PDFs merged successfully: {os.path.basename(output_file)}")
                messagebox.showinfo("Success", f"PDFs merged successfully:\n{output_file}")
            except Exception as e:
                messagebox.showerror("Error", f"Error merging PDFs: {str(e)}")
    
    def on_canvas_resize(self, event):
        if hasattr(self, 'doc'):
            self.render_pdf_page()

    # PDF Splitter methods
    def browse_pdf_to_split(self):
        file = filedialog.askopenfilename(
            title="Select PDF to Split",
            filetypes=[("PDF files", "*.pdf")]
        )
        
        if file:
            self.split_pdf_path.set(file)
            self.status_var.set(f"Selected PDF: {os.path.basename(file)}")
    
    def parse_page_ranges(self, range_str):
        pages = []
        parts = range_str.split(',')
        
        for part in parts:
            if '-' in part:
                start, end = part.split('-')
                try:
                    start, end = int(start.strip()), int(end.strip())
                    pages.extend(range(start, end + 1))
                except ValueError:
                    messagebox.showerror("Error", f"Invalid range format: {part}")
                    return []
            else:
                try:
                    pages.append(int(part.strip()))
                except ValueError:
                    messagebox.showerror("Error", f"Invalid page number: {part}")
                    return []
        
        return pages
    
    def split_pdf(self):
        pdf_path = self.split_pdf_path.get()
        
        if not pdf_path:
            messagebox.showwarning("Warning", "Please select a PDF file first")
            return
            
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                total_pages = len(pdf_reader.pages)
                
                if self.split_option.get() == "each":
                    # Split each page into a separate PDF
                    output_dir = filedialog.askdirectory(title="Select Output Directory")
                    if not output_dir:
                        return
                        
                    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
                    
                    for page_num in range(total_pages):
                        pdf_writer = PyPDF2.PdfWriter()
                        pdf_writer.add_page(pdf_reader.pages[page_num])
                        
                        output_filename = os.path.join(output_dir, f"{base_name}_page_{page_num+1}.pdf")
                        with open(output_filename, 'wb') as output_pdf:
                            pdf_writer.write(output_pdf)
                    
                    self.status_var.set(f"Split {total_pages} pages into separate PDFs")
                    messagebox.showinfo("Success", f"Split {total_pages} pages into separate PDFs in:\n{output_dir}")
                
                else:  # split by range
                    range_str = self.page_range.get()
                    if not range_str:
                        messagebox.showwarning("Warning", "Please enter a page range")
                        return
                        
                    pages = self.parse_page_ranges(range_str)
                    if not pages:
                        return
                        
                    # Filter valid pages
                    pages = [p for p in pages if 1 <= p <= total_pages]
                    
                    if not pages:
                        messagebox.showwarning("Warning", f"No valid pages in range. Total pages: {total_pages}")
                        return
                    
                    output_file = filedialog.asksaveasfilename(
                        title="Save Extracted Pages As",
                        defaultextension=".pdf",
                        filetypes=[("PDF files", "*.pdf")]
                    )
                    
                    if output_file:
                        pdf_writer = PyPDF2.PdfWriter()
                        
                        for page_num in pages:
                            pdf_writer.add_page(pdf_reader.pages[page_num-1])
                        
                        with open(output_file, 'wb') as output_pdf:
                            pdf_writer.write(output_pdf)
                        
                        self.status_var.set(f"Created PDF with {len(pages)} pages")
                        messagebox.showinfo("Success", f"Created PDF with {len(pages)} pages:\n{output_file}")
                    
        except Exception as e:
            messagebox.showerror("Error", f"Error splitting PDF: {str(e)}")
    
    # PDF Page Remover methods
    def browse_pdf_to_remove_pages(self):
        file = filedialog.askopenfilename(
            title="Select PDF",
            filetypes=[("PDF files", "*.pdf")]
        )
        
        if file:
            self.remove_pdf_path.set(file)
            try:
                with open(file, 'rb') as f:
                    pdf = PyPDF2.PdfReader(f)
                    total_pages = len(pdf.pages)
                    self.total_pages_var.set(f"Total Pages: {total_pages}")
            except Exception as e:
                messagebox.showerror("Error", f"Error opening PDF: {str(e)}")
    
    def remove_pdf_pages(self):
        pdf_path = self.remove_pdf_path.get()
        pages_to_remove_str = self.pages_to_remove.get()
        
        if not pdf_path:
            messagebox.showwarning("Warning", "Please select a PDF file first")
            return
            
        if not pages_to_remove_str:
            messagebox.showwarning("Warning", "Please specify pages to remove")
            return
        
        try:
            # Parse pages to remove
            pages_to_remove = self.parse_page_ranges(pages_to_remove_str)
            
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                total_pages = len(pdf_reader.pages)
                
                # Filter valid pages
                pages_to_remove = [p for p in pages_to_remove if 1 <= p <= total_pages]
                
                if not pages_to_remove:
                    messagebox.showwarning("Warning", f"No valid pages to remove. Total pages: {total_pages}")
                    return
                
                output_file = filedialog.asksaveasfilename(
                    title="Save Modified PDF As",
                    defaultextension=".pdf",
                    filetypes=[("PDF files", "*.pdf")]
                )
                
                if output_file:
                    pdf_writer = PyPDF2.PdfWriter()
                    
                    # Add all pages except those to be removed
                    for page_num in range(total_pages):
                        if page_num + 1 not in pages_to_remove:
                            pdf_writer.add_page(pdf_reader.pages[page_num])
                    
                    with open(output_file, 'wb') as output_pdf:
                        pdf_writer.write(output_pdf)
                    
                    self.status_var.set(f"Removed {len(pages_to_remove)} pages")
                    messagebox.showinfo("Success", f"Removed {len(pages_to_remove)} pages. New PDF saved as:\n{output_file}")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error removing pages: {str(e)}")
    
    # Fill and Sign methods
    def browse_pdf_to_sign(self):
        file = filedialog.askopenfilename(
            title="Select PDF",
            filetypes=[("PDF files", "*.pdf")]
        )
        
        if file:
            self.sign_pdf_path.set(file)
            self.status_var.set(f"Selected PDF: {os.path.basename(file)}")
    
    def add_signature_image(self):
        file = filedialog.askopenfilename(
            title="Select Signature Image",
            filetypes=[("Image files", "*.jpg *.jpeg *.png *.bmp")]
        )
        
        if file:
            self.signature_path.set(os.path.basename(file))
            self._signature_file = file
            self.status_var.set(f"Signature image added: {os.path.basename(file)}")
    
    def add_text_to_pdf(self):
        pdf_path = self.sign_pdf_path.get()
        text = self.field_text.get()
        
        if not pdf_path:
            messagebox.showwarning("Warning", "Please select a PDF file first")
            return
            
        if not text:
            messagebox.showwarning("Warning", "Please enter text to add")
            return
            
        try:
            page = int(self.field_page.get())
            x = int(self.field_x.get())
            y = int(self.field_y.get())
            
            # Create temporary file
            temp_output = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            temp_output.close()
            
            # Add text using PyMuPDF
            doc = fitz.open(pdf_path)
            
            if page < 1 or page > len(doc):
                messagebox.showwarning("Warning", f"Invalid page number. PDF has {len(doc)} pages.")
                return
                
            page_obj = doc[page-1]
            page_obj.insert_text((x, y), text, fontsize=12, color=(0, 0, 0))
            
            doc.save(temp_output.name)
            doc.close()
            
            # Update the PDF path to the modified file
            self.sign_pdf_path.set(temp_output.name)
            self.status_var.set(f"Text added to PDF at position ({x}, {y}) on page {page}")
            messagebox.showinfo("Success", "Text added to PDF")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error adding text: {str(e)}")
    
    def place_signature(self):
        pdf_path = self.sign_pdf_path.get()
        
        if not pdf_path:
            messagebox.showwarning("Warning", "Please select a PDF file first")
            return
            
        if not hasattr(self, '_signature_file'):
            messagebox.showwarning("Warning", "Please add a signature image first")
            return
            
        try:
            page = int(self.sig_page.get())
            x = int(self.sig_x.get())
            y = int(self.sig_y.get())
            width = int(self.sig_width.get())
            
            # Create temporary file
            temp_output = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            temp_output.close()
            
            # Add signature using PyMuPDF
            doc = fitz.open(pdf_path)
            
            if page < 1 or page > len(doc):
                messagebox.showwarning("Warning", f"Invalid page number. PDF has {len(doc)} pages.")
                return
                
            page_obj = doc[page-1]
            
            # Calculate height based on aspect ratio
            img = Image.open(self._signature_file)
            aspect_ratio = img.height / img.width
            height = int(width * aspect_ratio)
            
            # Add the image to the PDF
            img_rect = fitz.Rect(x, y, x + width, y + height)
            page_obj.insert_image(img_rect, filename=self._signature_file)
            
            doc.save(temp_output.name)
            doc.close()
            
            # Update the PDF path to the modified file
            self.sign_pdf_path.set(temp_output.name)
            self.status_var.set(f"Signature placed on PDF at position ({x}, {y}) on page {page}")
            messagebox.showinfo("Success", "Signature placed on PDF")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error placing signature: {str(e)}")
    
    def save_filled_pdf(self):
        if not hasattr(self, 'doc'):
            messagebox.showwarning("Warning", "No PDF to save")
            return
        
        output_file = filedialog.asksaveasfilename(
            title="Save Filled PDF As",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")]
        )
    
        if output_file:
            try:
                self.doc.save(output_file)
                self.status_var.set(f"PDF saved successfully: {os.path.basename(output_file)}")
                messagebox.showinfo("Success", f"PDF saved successfully:\n{output_file}")
            except Exception as e:
                messagebox.showerror("Error", f"Error saving PDF: {str(e)}")


def main():
    root = tk.Tk()
    root.title("PDF Toolkit")
    
    # Set app icon if available
    try:
        root.iconbitmap("pdf_icon.ico")  # You would need to create this icon file
    except:
        pass
    
    # Make DPI aware for better display on high-resolution screens
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass
    
    PDFToolsApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()