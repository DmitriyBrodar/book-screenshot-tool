import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pyautogui
import time
import os
import subprocess
import platform
from PIL import Image, ImageTk
import threading
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import glob

class RegionSelector:
    """Handles selecting what part of the screen to capture"""
    def __init__(self, callback):
        self.callback = callback
        self.start_x = self.start_y = 0
        self.end_x = self.end_y = 0
        self.selecting = False
        
    def start_selection(self):
        # Create a semi-transparent overlay window covering the whole screen
        self.root = tk.Toplevel()
        self.root.attributes('-fullscreen', True)
        self.root.attributes('-alpha', 0.3)  # Make it see-through
        self.root.configure(bg='black')
        self.root.attributes('-topmost', True)  # Keep it on top
        self.root.focus_force()
        
        # Create the drawing area
        self.canvas = tk.Canvas(self.root, highlightthickness=0, bg='black')
        self.canvas.pack(fill=tk.BOTH, expand=True)
        
        # Show instructions at the top
        self.canvas.create_text(
            self.root.winfo_screenwidth()//2, 50,
            text="Click and drag to select capture area • Press ESC to cancel",
            fill='white', font=('Segoe UI', 16, 'bold')
        )
        
        # Set up mouse events for drawing the selection box
        self.canvas.bind('<Button-1>', self.on_click)
        self.canvas.bind('<B1-Motion>', self.on_drag)
        self.canvas.bind('<ButtonRelease-1>', self.on_release)
        self.root.bind('<Escape>', lambda e: self.cancel())
        
        self.canvas.focus_set()
        
    def on_click(self, event):
        # Remember where the user started clicking
        self.start_x, self.start_y = event.x, event.y
        self.selecting = True
        
    def on_drag(self, event):
        # Draw a red rectangle as the user drags the mouse
        if self.selecting:
            self.canvas.delete("selection")  # Remove old rectangle
            self.end_x, self.end_y = event.x, event.y
            self.canvas.create_rectangle(
                self.start_x, self.start_y, self.end_x, self.end_y,
                outline='red', width=3, fill='red', stipple='gray50',
                tags="selection"
            )
            
    def on_release(self, event):
        # When user lets go of mouse, finalize the selection
        if self.selecting:
            self.selecting = False
            # Make sure we have the corners in the right order
            x1, y1 = min(self.start_x, self.end_x), min(self.start_y, self.end_y)
            x2, y2 = max(self.start_x, self.end_x), max(self.start_y, self.end_y)
            
            # Only accept the selection if it's big enough
            if abs(x2 - x1) > 10 and abs(y2 - y1) > 10:
                # Convert to the format pyautogui needs: (x, y, width, height)
                region = (x1, y1, x2 - x1, y2 - y1)
                self.root.destroy()
                self.callback(region)
            else:
                self.cancel()
    
    def cancel(self):
        # Close the selection window and return nothing
        self.root.destroy()
        self.callback(None)

class ModernBookScreenshotTool:
    def __init__(self, root):
        self.root = root
        self.root.title("📖 Book Screenshot Tool")
        self.root.geometry("850x830")
        self.root.configure(bg='#f8f9fa')
        self.root.resizable(True, True)
        self.root.minsize(800, 600)
        
        # Set up all the variables we'll need
        self.is_running = False
        self.save_folder = ""
        self.screenshot_count = 0
        self.region = None  # The area of screen to capture
        self.screenshots = []  # List of screenshot file paths
        self.click_position = None  # Where to click for page turning
        self.countdown = 10
        
        # Define all the colors used in the interface
        self.colors = {
            'bg': '#f8f9fa',
            'surface': '#ffffff',
            'primary': '#2563eb',
            'primary_hover': '#1d4ed8',
            'success': '#16a34a',
            'success_hover': '#15803d',
            'danger': '#dc2626',
            'danger_hover': '#b91c1c',
            'text': '#1f2937',
            'text_muted': '#6b7280',
            'border': '#e5e7eb'
        }
        
        # Safety feature - move mouse to corner to stop the program
        pyautogui.FAILSAFE = True
        self.create_ui()
    
    def create_ui(self):
        # Create the top header with title
        header = tk.Frame(self.root, bg=self.colors['bg'], height=80)
        header.pack(fill=tk.X, padx=20, pady=(20, 0))
        header.pack_propagate(False)  # Don't let it shrink
        
        title = tk.Label(header, text="📖 Book Screenshot Tool", 
                        font=('Segoe UI', 20, 'bold'),
                        bg=self.colors['bg'], fg=self.colors['text'])
        title.pack(side=tk.LEFT, pady=20)
        
        subtitle = tk.Label(header, text="Capture book pages and create PDFs effortlessly",
                           font=('Segoe UI', 11),
                           bg=self.colors['bg'], fg=self.colors['text_muted'])
        subtitle.pack(side=tk.LEFT, padx=(15, 0), pady=25)
        
        # Create main content area
        main = tk.Frame(self.root, bg=self.colors['bg'])
        main.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Split into left and right columns
        left_frame = tk.Frame(main, bg=self.colors['bg'])
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        # Left side has setup options and page turning method
        self.create_setup_section(left_frame)
        self.create_method_section(left_frame)
        
        # Right side has controls and status
        right_frame = tk.Frame(main, bg=self.colors['bg'])
        right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 0))
        
        self.create_control_section(right_frame)
    
    def create_card(self, parent, title, icon=""):
        """Creates a white card with a title - used for organizing the interface"""
        card = tk.Frame(parent, bg=self.colors['surface'], relief='flat')
        card.pack(fill=tk.X, pady=(0, 20))
        
        # Card header with title
        header = tk.Frame(card, bg=self.colors['surface'])
        header.pack(fill=tk.X, padx=25, pady=(20, 15))
        
        title_label = tk.Label(header, text=f"{icon} {title}",
                              font=('Segoe UI', 12, 'bold'),
                              bg=self.colors['surface'], fg=self.colors['text'])
        title_label.pack(side=tk.LEFT)
        
        # Content area inside the card
        content = tk.Frame(card, bg=self.colors['surface'])
        content.pack(fill=tk.BOTH, expand=True, padx=25, pady=(0, 20))
        
        return content
    
    def create_setup_section(self, parent):
        """Creates the setup card with folder selection, region selection, and settings"""
        content = self.create_card(parent, "Setup", "⚙️")
        
        # Folder selection area
        folder_frame = tk.Frame(content, bg=self.colors['surface'])
        folder_frame.pack(fill=tk.X, pady=(0, 15))
        
        tk.Label(folder_frame, text="Save Location:", font=('Segoe UI', 9, 'bold'),
                bg=self.colors['surface'], fg=self.colors['text']).pack(anchor=tk.W, pady=(0, 5))
        
        folder_input = tk.Frame(folder_frame, bg=self.colors['surface'])
        folder_input.pack(fill=tk.X)
        
        # Text field showing selected folder (read-only)
        self.folder_var = tk.StringVar(value="Click Browse to select folder...")
        folder_entry = tk.Entry(folder_input, textvariable=self.folder_var,
                               font=('Segoe UI', 9), state='readonly',
                               bg='#f9fafb', fg=self.colors['text_muted'],
                               relief='solid', bd=1)
        folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        # Browse button to select folder
        browse_btn = tk.Button(folder_input, text="Browse", font=('Segoe UI', 9),
                              bg=self.colors['primary'], fg='white', relief='flat',
                              padx=20, pady=8, cursor='hand2',
                              command=self.browse_folder)
        browse_btn.pack(side=tk.RIGHT)
        
        # Screen region selection area
        region_frame = tk.Frame(content, bg=self.colors['surface'])
        region_frame.pack(fill=tk.X, pady=(0, 15))
        
        tk.Label(region_frame, text="Capture Area:", font=('Segoe UI', 9, 'bold'),
                bg=self.colors['surface'], fg=self.colors['text']).pack(anchor=tk.W, pady=(0, 5))
        
        # Button to start region selection
        self.region_var = tk.StringVar(value="Click to select screen area...")
        region_btn = tk.Button(region_frame, textvariable=self.region_var,
                              font=('Segoe UI', 9), bg='#f9fafb', fg=self.colors['text_muted'],
                              relief='solid', bd=1, padx=15, pady=10, cursor='hand2',
                              command=self.select_region, anchor=tk.W)
        region_btn.pack(fill=tk.X)
        
        # Settings area
        settings_frame = tk.Frame(content, bg=self.colors['surface'])
        settings_frame.pack(fill=tk.X)
        
        # Put pages and delay inputs side by side
        inputs_frame = tk.Frame(settings_frame, bg=self.colors['surface'])
        inputs_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Number of pages to capture
        pages_frame = tk.Frame(inputs_frame, bg=self.colors['surface'])
        pages_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        tk.Label(pages_frame, text="Pages:", font=('Segoe UI', 9, 'bold'),
                bg=self.colors['surface'], fg=self.colors['text']).pack(anchor=tk.W, pady=(0, 2))
        
        self.pages_var = tk.StringVar(value="10")
        pages_entry = tk.Entry(pages_frame, textvariable=self.pages_var,
                              font=('Segoe UI', 9), bg=self.colors['surface'],
                              relief='solid', bd=1, width=8)
        pages_entry.pack(anchor=tk.W)
        
        # Delay between screenshots
        delay_frame = tk.Frame(inputs_frame, bg=self.colors['surface'])
        delay_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        tk.Label(delay_frame, text="Delay (sec):", font=('Segoe UI', 9, 'bold'),
                bg=self.colors['surface'], fg=self.colors['text']).pack(anchor=tk.W, pady=(0, 2))
        
        self.delay_var = tk.StringVar(value="2")
        delay_entry = tk.Entry(delay_frame, textvariable=self.delay_var,
                              font=('Segoe UI', 9), bg=self.colors['surface'],
                              relief='solid', bd=1, width=8)
        delay_entry.pack(anchor=tk.W)
        
        # PDF creation options
        pdf_frame = tk.Frame(settings_frame, bg=self.colors['surface'])
        pdf_frame.pack(fill=tk.X)
        
        # Checkbox to automatically create PDF
        self.create_pdf_var = tk.BooleanVar(value=True)
        pdf_check = tk.Checkbutton(pdf_frame, text="Auto-create PDF",
                                  variable=self.create_pdf_var,
                                  font=('Segoe UI', 9, 'bold'),
                                  bg=self.colors['surface'], fg=self.colors['text'])
        pdf_check.pack(anchor=tk.W, pady=(0, 5))
        
        # Text field for PDF filename
        self.pdf_name_var = tk.StringVar(value="book_screenshots.pdf")
        pdf_entry = tk.Entry(pdf_frame, textvariable=self.pdf_name_var,
                            font=('Segoe UI', 9), bg=self.colors['surface'],
                            relief='solid', bd=1)
        pdf_entry.pack(fill=tk.X)
    
    def create_method_section(self, parent):
        """Creates the page turning method card"""
        content = self.create_card(parent, "Page Turning Method", "🔄")
        
        # Variable to track which method is selected
        self.method_var = tk.StringVar(value="keyboard")
        
        # Radio buttons to choose method
        method_frame = tk.Frame(content, bg=self.colors['surface'])
        method_frame.pack(fill=tk.X, pady=(0, 15))
        
        keyboard_radio = tk.Radiobutton(method_frame, text="Keyboard Key",
                                       variable=self.method_var, value="keyboard",
                                       font=('Segoe UI', 9, 'bold'),
                                       bg=self.colors['surface'], fg=self.colors['text'])
        keyboard_radio.pack(anchor=tk.W, pady=(0, 5))
        
        mouse_radio = tk.Radiobutton(method_frame, text="Mouse Click",
                                    variable=self.method_var, value="mouse",
                                    font=('Segoe UI', 9, 'bold'),
                                    bg=self.colors['surface'], fg=self.colors['text'])
        mouse_radio.pack(anchor=tk.W)
        
        # Options for each method
        options_frame = tk.Frame(content, bg=self.colors['surface'])
        options_frame.pack(fill=tk.X)
        
        # Dropdown to select which key to press
        key_frame = tk.Frame(options_frame, bg=self.colors['surface'])
        key_frame.pack(fill=tk.X, pady=(0, 10))
        
        tk.Label(key_frame, text="Key:", font=('Segoe UI', 9),
                bg=self.colors['surface'], fg=self.colors['text']).pack(side=tk.LEFT, padx=(0, 10))
        
        self.key_var = tk.StringVar(value="right")
        key_combo = ttk.Combobox(key_frame, textvariable=self.key_var,
                                values=['right', 'left', 'space', 'pagedown', 'pageup', 'enter'],
                                font=('Segoe UI', 9), width=12, state='readonly')
        key_combo.pack(side=tk.LEFT)
        
        # Button to set where to click for mouse method
        self.click_var = tk.StringVar(value="Click to set position...")
        click_btn = tk.Button(options_frame, textvariable=self.click_var,
                             font=('Segoe UI', 9), bg='#f9fafb', fg=self.colors['text_muted'],
                             relief='solid', bd=1, padx=15, pady=8, cursor='hand2',
                             command=self.set_click_position, anchor=tk.W)
        click_btn.pack(fill=tk.X)
    
    def create_control_section(self, parent):
        """Creates the controls card with buttons and status"""
        content = self.create_card(parent, "Controls", "🎮")
        
        # Big green start button
        self.start_button = tk.Button(content, text="🚀 Start Capture",
                                     font=('Segoe UI', 10, 'bold'),
                                     bg=self.colors['success'], fg='white',
                                     relief='flat', padx=30, pady=15, cursor='hand2',
                                     command=self.start_screenshot)
        self.start_button.pack(fill=tk.X, pady=(0, 10))
        
        # Stop button (disabled by default)
        self.stop_button = tk.Button(content, text="⏹️ Stop Capture",
                                    font=('Segoe UI', 10, 'bold'),
                                    bg='white', fg='white',
                                    relief='solid', bd=2, padx=30, pady=15, cursor='hand2',
                                    command=self.stop_screenshot, state=tk.DISABLED)
        self.stop_button.pack(fill=tk.X, pady=(0, 15))
        
        # Status display area
        status_frame = tk.Frame(content, bg=self.colors['surface'])
        status_frame.pack(fill=tk.X, pady=(0, 15))
        
        tk.Label(status_frame, text="Status:", font=('Segoe UI', 9, 'bold'),
                bg=self.colors['surface'], fg=self.colors['text']).pack(anchor=tk.W)
        
        # Text that shows what the program is currently doing
        self.status_var = tk.StringVar(value="Ready to capture")
        status_label = tk.Label(status_frame, textvariable=self.status_var,
                               font=('Segoe UI', 9), bg=self.colors['surface'],
                               fg=self.colors['text_muted'], wraplength=200)
        status_label.pack(anchor=tk.W, pady=(2, 0))
        
        # Progress bar area
        progress_frame = tk.Frame(content, bg=self.colors['surface'])
        progress_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Visual progress bar
        self.progress = ttk.Progressbar(progress_frame, length=200, mode='determinate')
        self.progress.pack(fill=tk.X, pady=(0, 5))
        
        # Text showing progress numbers
        self.progress_text_var = tk.StringVar(value="0 / 0 pages")
        progress_label = tk.Label(progress_frame, textvariable=self.progress_text_var,
                                 font=('Segoe UI', 8), bg=self.colors['surface'],
                                 fg=self.colors['text_muted'])
        progress_label.pack(anchor=tk.W)
        
        # Button to create PDF from images that already exist
        pdf_btn = tk.Button(content, text="📄 Create PDF from Images",
                           font=('Segoe UI', 9), bg=self.colors['surface'],
                           fg=self.colors['primary'], relief='solid', bd=1,
                           padx=20, pady=10, cursor='hand2',
                           command=self.create_pdf_from_existing)
        pdf_btn.pack(fill=tk.X)
    
    def browse_folder(self):
        """Opens a dialog to let user pick where to save files"""
        folder = filedialog.askdirectory()
        if folder:
            self.save_folder = folder
            # Show just the folder name, not the full path
            display_path = os.path.basename(folder) if folder else "..."
            self.folder_var.set(f"📁 {display_path}")

    def select_region(self):
        """Opens the region selector to let user choose what part of screen to capture"""
        def on_region_selected(region):
            if region:
                self.region = region
                # Show the dimensions of the selected area
                self.region_var.set(f"✅ {region[2]} × {region[3]} pixels")
            else:
                self.region_var.set("❌ Selection cancelled")
        
        # Create and start the region selector
        selector = RegionSelector(on_region_selected)
        selector.start_selection()
    
    def set_click_position(self):
        """Let user set where to click for turning pages with mouse method"""
        # Hide the main window so user can see what's behind it
        self.root.withdraw()
        result = messagebox.askokcancel("Set Click Position", 
                                       "Position your mouse on the 'Next Page' button and click OK")
        if result:
            time.sleep(2)  # Give user time to position their mouse
            x, y = pyautogui.position()  # Get current mouse position
            self.click_position = (x, y)
            self.click_var.set(f"✅ Position: {x}, {y}")
        else:
            self.click_var.set("❌ Position not set")
        # Show the main window again
        self.root.deiconify()
    
    def start_screenshot(self):
        """Main function that starts the screenshot process after checking everything is ready"""
        # Check if user has set up everything needed
        if not self.save_folder:
            messagebox.showerror("Error", "Please select a save folder")
            return
        
        if not self.region:
            messagebox.showerror("Error", "Please select screenshot region")
            return
        
        if self.method_var.get() == "mouse" and not self.click_position:
            messagebox.showerror("Error", "Please set click position for mouse method")
            return
        
        # Check if the numbers entered are valid
        try:
            pages = int(self.pages_var.get())
            delay = float(self.delay_var.get())
            if pages <= 0 or delay < 0:
                raise ValueError()
        except ValueError:
            messagebox.showerror("Error", "Please enter valid numbers")
            return
        
        # Change the UI to show we're running
        self.is_running = True
        self.start_button.config(state=tk.DISABLED, bg='#d1d5db')
        self.stop_button.config(state=tk.NORMAL, bg=self.colors['danger'])
        self.progress['maximum'] = pages
        self.progress['value'] = 0
        self.screenshot_count = 0
        self.screenshots = []
        
        # Start the actual work in a separate thread so UI doesn't freeze
        thread = threading.Thread(target=self.screenshot_process, args=(pages, delay))
        thread.daemon = True  # Thread will close when main program closes
        thread.start()
    
    def screenshot_process(self, pages, delay):
        """This is the main work function that takes all the screenshots"""
        # Give user 3 seconds to get ready
        for i in range(3, 0, -1):
            if not self.is_running:  # User clicked stop
                return
            self.status_var.set(f"⏱️ Starting in {i} seconds...")
            time.sleep(1)
        
        # Take screenshots of each page
        for page in range(pages):
            if not self.is_running:  # User clicked stop
                break
            
            # Update status to show current progress
            self.status_var.set(f"📸 Capturing page {page + 1}...")
            self.progress_text_var.set(f"{page + 1} / {pages} pages")
            
            # Take the actual screenshot
            screenshot = pyautogui.screenshot(region=self.region)
            filename = f"page_{page + 1:03d}.png"  # page_001.png, page_002.png, etc.
            filepath = os.path.join(self.save_folder, filename)
            screenshot.save(filepath)
            
            # Keep track of what we've done
            self.screenshots.append(filepath)
            self.screenshot_count += 1
            self.progress['value'] = self.screenshot_count
            self.root.update_idletasks()  # Refresh the UI
            
            # Turn to next page (except on the last page)
            if page < pages - 1:
                if self.method_var.get() == "keyboard":
                    pyautogui.press(self.key_var.get())
                else:
                    pyautogui.click(self.click_position[0], self.click_position[1])
                time.sleep(delay)  # Wait before taking next screenshot
        
        # Create PDF if user wants it
        pdf_created = False
        pdf_name = ""
        
        if self.create_pdf_var.get() and self.screenshots:
            self.status_var.set("📄 Creating PDF...")
            self.root.update_idletasks()
            pdf_created, pdf_name = self.create_pdf(self.screenshots)
        
        # Show the completion message
        self.show_completion_dialog(self.screenshot_count, pdf_created, pdf_name)
        
        # Put the UI back to normal
        self.is_running = False
        self.start_button.config(state=tk.NORMAL, bg=self.colors['success'])
        self.stop_button.config(state=tk.DISABLED, bg='#f3f4f6')
    
    def stop_screenshot(self):
        """Stop the screenshot process when user clicks stop button"""
        self.is_running = False
        self.start_button.config(state=tk.NORMAL, bg=self.colors['success'])
        self.stop_button.config(state=tk.DISABLED, bg='#f3f4f6')
        self.status_var.set("⏹️ Stopped by user")
    
    def create_pdf(self, image_paths):
        """Convert all the screenshot images into one PDF file"""
        try:
            # Make sure filename ends with .pdf
            pdf_filename = self.pdf_name_var.get()
            if not pdf_filename.endswith('.pdf'):
                pdf_filename += '.pdf'
            
            # Create the PDF
            pdf_path = os.path.join(self.save_folder, pdf_filename)
            c = canvas.Canvas(pdf_path, pagesize=A4)
            page_width, page_height = A4
            
            # Add each image as a page in the PDF
            for i, image_path in enumerate(image_paths):
                if not os.path.exists(image_path):
                    continue
                
                self.status_var.set(f"📄 Adding page {i+1}/{len(image_paths)} to PDF...")
                self.root.update_idletasks()
                
                img = Image.open(image_path)
                img_width, img_height = img.size
                
                # Figure out how much to shrink/expand the image to fit the page
                scale_x = page_width / img_width
                scale_y = page_height / img_height
                scale = min(scale_x, scale_y)  # Use smaller scale so image fits
                
                # Center the image on the page
                scaled_width = img_width * scale
                scaled_height = img_height * scale
                x = (page_width - scaled_width) / 2
                y = (page_height - scaled_height) / 2
                
                c.drawImage(image_path, x, y, width=scaled_width, height=scaled_height)
                c.showPage()  # Move to next page
            
            c.save()  # Finish and save the PDF
            self.status_var.set(f"✅ PDF created: {pdf_filename}")
            return True, pdf_filename
            
        except Exception as e:
            messagebox.showerror("PDF Error", f"Error creating PDF: {str(e)}")
            return False, ""
    
    def create_pdf_from_existing(self):
        """Create PDF from images that are already saved in the folder"""
        if not self.save_folder:
            messagebox.showerror("Error", "Please select a folder first")
            return
        
        # Look for all types of image files
        image_extensions = ['*.png', '*.jpg', '*.jpeg', '*.bmp', '*.tiff']
        image_files = []
        
        for ext in image_extensions:
            image_files.extend(glob.glob(os.path.join(self.save_folder, ext)))
        
        if not image_files:
            messagebox.showwarning("No Images", "No image files found in the selected folder")
            return
        
        # Put files in order (page_001.png comes before page_002.png)
        image_files.sort()
        
        # Ask user what to name the PDF
        pdf_name = simpledialog.askstring("PDF Name", 
                                        "Enter PDF filename:", 
                                        initialvalue="book_from_images.pdf")
        
        if not pdf_name:
            return
        
        if not pdf_name.endswith('.pdf'):
            pdf_name += '.pdf'
        
        try:
            self.status_var.set("📄 Creating PDF from existing images...")
            self.root.update_idletasks()
            
            # Same PDF creation process as before
            pdf_path = os.path.join(self.save_folder, pdf_name)
            c = canvas.Canvas(pdf_path, pagesize=A4)
            page_width, page_height = A4
            
            for i, image_path in enumerate(image_files):
                self.status_var.set(f"📄 Adding page {i+1}/{len(image_files)} to PDF...")
                self.root.update_idletasks()
                
                try:
                    img = Image.open(image_path)
                    img_width, img_height = img.size
                    
                    # Scale and center image on PDF page
                    scale_x = page_width / img_width
                    scale_y = page_height / img_height
                    scale = min(scale_x, scale_y)
                    
                    scaled_width = img_width * scale
                    scaled_height = img_height * scale
                    x = (page_width - scaled_width) / 2
                    y = (page_height - scaled_height) / 2
                    
                    c.drawImage(image_path, x, y, width=scaled_width, height=scaled_height)
                    c.showPage()
                    
                except Exception as e:
                    print(f"Error processing {image_path}: {e}")
                    continue  # Skip this image and try the next one
            
            c.save()
            
            # Tell user it worked
            self.status_var.set(f"✅ PDF created: {pdf_name}")
            messagebox.showinfo("Success", 
                               f"PDF created successfully!\n\n"
                               f"File: {pdf_name}\n"
                               f"Pages: {len(image_files)}\n"
                               f"Location: {self.save_folder}")
            
        except Exception as e:
            messagebox.showerror("PDF Error", f"Error creating PDF: {str(e)}")
            self.status_var.set("❌ PDF creation failed")
    
    def open_folder(self, folder_path):
        """Open the save folder in Windows Explorer, Mac Finder, or Linux file manager"""
        try:
            if platform.system() == "Windows":
                os.startfile(folder_path)
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", folder_path])
            else:  # Linux
                subprocess.run(["xdg-open", folder_path])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open folder: {str(e)}")
    
    def show_completion_dialog(self, screenshots_count, pdf_created=False, pdf_name=""):
        """Show a nice dialog when everything is finished"""
        # Create popup window
        dialog = tk.Toplevel(self.root)
        dialog.title("✅ Capture Complete")
        dialog.geometry("400x300")
        dialog.configure(bg=self.colors['surface'])
        dialog.resizable(False, False)
        dialog.transient(self.root)  # Keep it connected to main window
        dialog.grab_set()  # Make it modal (user must deal with it first)
        
        # Put dialog in center of screen
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (400 // 2)
        y = (dialog.winfo_screenheight() // 2) - (300 // 2)
        dialog.geometry(f"400x300+{x}+{y}")
        
        # Green header section
        header_frame = tk.Frame(dialog, bg=self.colors['success'], height=60)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        success_label = tk.Label(header_frame, text="🎉 Capture Completed Successfully!",
                                font=('Segoe UI', 14, 'bold'),
                                bg=self.colors['success'], fg='white')
        success_label.pack(expand=True)
        
        # Main content area
        content_frame = tk.Frame(dialog, bg=self.colors['surface'])
        content_frame.pack(fill=tk.BOTH, expand=True, padx=25, pady=20)
        
        # Summary of what was done
        summary_frame = tk.Frame(content_frame, bg=self.colors['surface'])
        summary_frame.pack(fill=tk.X, pady=(0, 20))
        
        tk.Label(summary_frame, text="📊 Summary:",
                font=('Segoe UI', 11, 'bold'),
                bg=self.colors['surface'], fg=self.colors['text']).pack(anchor=tk.W)
        
        # Build summary text
        summary_text = f"• {screenshots_count} screenshots captured\n"
        summary_text += f"• Saved to: {os.path.basename(self.save_folder)}\n"
        
        if pdf_created:
            summary_text += f"• PDF created: {pdf_name}"
        else:
            summary_text += "• No PDF created"
        
        tk.Label(summary_frame, text=summary_text,
                font=('Segoe UI', 9),
                bg=self.colors['surface'], fg=self.colors['text_muted'],
                justify=tk.LEFT).pack(anchor=tk.W, pady=(5, 0))
        
        # Buttons at bottom
        button_frame = tk.Frame(content_frame, bg=self.colors['surface'])
        button_frame.pack(fill=tk.X, pady=(20, 0))
        
        # Button to open the folder where files were saved
        open_btn = tk.Button(button_frame, text="📁 Open Folder",
                            font=('Segoe UI', 10, 'bold'),
                            bg=self.colors['primary'], fg='white',
                            relief='flat', padx=25, pady=10, cursor='hand2',
                            command=lambda: [self.open_folder(self.save_folder), dialog.destroy()])
        open_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Button to close the dialog
        close_btn = tk.Button(button_frame, text="✓ Close",
                             font=('Segoe UI', 10),
                             bg=self.colors['surface'], fg=self.colors['text'],
                             relief='solid', bd=1, padx=25, pady=10, cursor='hand2',
                             command=dialog.destroy)
        close_btn.pack(side=tk.RIGHT)
        
        # This auto-close code is commented out but could be used
        """
        # Auto-close timer (optional)
        timer_frame = tk.Frame(content_frame, bg=self.colors['surface'])
        timer_frame.pack(fill=tk.X, pady=(15, 0))
        
        self.timer_var = tk.StringVar(value="Auto-close in 10 seconds...")
        timer_label = tk.Label(timer_frame, textvariable=self.timer_var,
                              font=('Segoe UI', 8),
                              bg=self.colors['surface'], fg=self.colors['text_muted'])
        timer_label.pack()
        
        # Countdown timer
        self.countdown = 60
        self.countdown_timer(dialog)
        """
        
        dialog.focus_set()  # Make sure dialog gets keyboard focus
        return dialog
    
    def countdown_timer(self, dialog):
        """Auto-close timer function (currently not used)"""
        if self.countdown > 0:
            self.timer_var.set(f"Auto-close in {self.countdown} seconds...")
            self.countdown -= 1
            dialog.after(1000, lambda: self.countdown_timer(dialog))  # Call again in 1 second
        else:
            if dialog.winfo_exists():  # Make sure dialog still exists
                dialog.destroy()

# This runs when the script is started directly (not imported)
if __name__ == "__main__":
    root = tk.Tk()
    app = ModernBookScreenshotTool(root)
    root.mainloop()  # Start the GUI