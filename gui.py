import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import webbrowser
from PIL import Image
from receipt_processing import extract_receipt_text_to_json
from route_sheet import update_route_sheet_from_json
import win32com.client

class RouteSheetApp:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("Route Sheet Assistant")
        self.root.geometry("1000x800")
        
        # Set appearance mode and theme
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")
        
        self.selected_files = []
        self.generated_files = []
        
        # Configure grid layout
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)
        
        self.create_widgets()
        
    def create_widgets(self):
        # Main container with padding
        self.main_container = ctk.CTkFrame(self.root, corner_radius=15)
        self.main_container.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.main_container.grid_columnconfigure(0, weight=1)

        # Logo section
        logo_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        logo_frame.grid(row=0, column=0, padx=20, pady=(20, 0), sticky="ew")
        
        # Load and display logo
        try:
            # Adjust the path to where your logo is stored
            logo_path = "images/logo.png"  
            logo_image = Image.open(logo_path)
            # Resize image if needed (adjust dimensions as needed)
            logo_image = logo_image.resize((250, 350))  
            logo_photo = ctk.CTkImage(light_image=logo_image, dark_image=logo_image, size=(200, 150))
            logo_label = ctk.CTkLabel(logo_frame, image=logo_photo, text="")
            logo_label.pack(pady=10, fill="x", expand=True)
        except Exception as e:
            print(f"Could not load logo: {e}")
        
        # Header section with modern styling
        header_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        header_frame.grid(row=1, column=0, padx=20, pady=(10, 10), sticky="ew")
        
        header_label = ctk.CTkLabel(
            header_frame,
            text="Route Sheet Assistant",
            font=ctk.CTkFont(size=32, weight="bold"),
        )
        header_label.pack(pady=10)
        
        subtitle = ctk.CTkLabel(
            header_frame,
            text="Upload Blackpug Receipts (PDF, JPG, PNG), Click Process, Download or Print Route Sheet",
            font=ctk.CTkFont(size=16),
            text_color="gray70"
        )
        subtitle.pack()
        
        # Action buttons container
        self.button_container = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.button_container.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        self.button_container.grid_columnconfigure((0, 1, 2), weight=1)
        
        # Modern styled buttons with icons (using Unicode characters as simple icons)
        self.select_button = ctk.CTkButton(
            self.button_container,
            text="📂 Select Receipts",
            command=self.select_receipts,
            height=45,
            font=ctk.CTkFont(size=15),
            fg_color="#2D5AF0",
            hover_color="#1E3EBF"
        )
        self.select_button.grid(row=0, column=0, padx=10, sticky="ew")
        
        self.process_button = ctk.CTkButton(
            self.button_container,
            text="⚙️ Process Receipts",
            command=self.process_receipts,
            height=45,
            font=ctk.CTkFont(size=15),
            fg_color="#14B8A6",
            hover_color="#0E8A7D"
        )
        self.process_button.grid(row=0, column=1, padx=10, sticky="ew")
        
        self.print_button = ctk.CTkButton(
            self.button_container,
            text="🖨️ Print Route Sheets",
            command=self.print_route_sheets,
            height=45,
            font=ctk.CTkFont(size=15),
            state="disabled",
            fg_color="#6366F1",
            hover_color="#4F46E5"
        )
        self.print_button.grid(row=0, column=2, padx=10, sticky="ew")
        
        # Status section with modern styling
        self.status_frame = ctk.CTkFrame(self.main_container, height=40, fg_color="#F8FAFC")
        self.status_frame.grid(row=3, column=0, padx=20, pady=10, sticky="ew")
        
        self.status_label = ctk.CTkLabel(
            self.status_frame,
            text="Ready to process receipts",
            font=ctk.CTkFont(size=14),
            text_color="gray60"
        )
        self.status_label.pack(pady=10)
        
        # Create tabview for data display and generated files
        self.tabview = ctk.CTkTabview(self.main_container)
        self.tabview.grid(row=3, column=0, padx=20, pady=10, sticky="nsew")
        self.main_container.grid_rowconfigure(4, weight=1)
        
        # Add tabs
        self.tabview.add("Processed Data")
        self.tabview.add("Generated Files")
        
        # Processed data display
        self.data_display = ctk.CTkTextbox(
            self.tabview.tab("Processed Data"),
            font=ctk.CTkFont(family="Courier", size=12)
        )
        self.data_display.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Generated files list
        self.files_frame = ctk.CTkScrollableFrame(
            self.tabview.tab("Generated Files"),
            label_text="Generated Route Sheets"
        )
        self.files_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
    def select_receipts(self):
        file_paths = filedialog.askopenfilenames(
            title="Select Receipt Documents",
            filetypes=[("Image and PDF Files", "*.png;*.jpg;*.jpeg;*.tiff;*.bmp;*.pdf")]
        )
        
        if not file_paths:
            return
            
        self.selected_files = file_paths
        self.status_label.configure(
            text=f"✅ {len(file_paths)} file(s) selected",
            text_color="#059669"
        )
        self.process_button.configure(state="normal")
        
    def process_receipts(self):
        self.generated_files = []
        self.data_display.delete("1.0", "end")
        
        try:
            self.status_label.configure(
                text="⏳ Processing receipts...",
                text_color="#2563EB"
            )
            self.root.update()
            
            for file_path in self.selected_files:
                receipt_data = extract_receipt_text_to_json(file_path)
                output_path = update_route_sheet_from_json(receipt_data)
                self.generated_files.append(output_path)
                
                # Display receipt data
                self.data_display.insert("end", f"\n=== {os.path.basename(file_path)} ===\n")
                formatted_data = self.format_receipt_data(receipt_data)
                self.data_display.insert("end", formatted_data + "\n")
                
                # Add file button to generated files tab
                self.add_file_button(output_path)
            
            self.status_label.configure(
                text="✅ All receipts processed successfully",
                text_color="#059669"
            )
            self.print_button.configure(state="normal")
            
        except Exception as e:
            self.status_label.configure(
                text="❌ Error processing receipts",
                text_color="#DC2626"
            )
            messagebox.showerror("Error", f"Failed to process receipts: {str(e)}")
            
    def add_file_button(self, file_path):
        """Add a styled button for each generated file"""
        button_frame = ctk.CTkFrame(self.files_frame, fg_color="transparent")
        button_frame.pack(fill="x", pady=5)
        
        file_button = ctk.CTkButton(
            button_frame,
            text=f"📄 {os.path.basename(file_path)}",
            command=lambda: webbrowser.open(file_path),
            fg_color="#3B82F6",
            hover_color="#2563EB",
            height=35
        )
        file_button.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
    def format_receipt_data(self, data):
        """Format receipt data for display"""
        formatted_lines = []
        for key, value in data.items():
            formatted_lines.append(f"{key.ljust(20)}: {value}")
        return "\n".join(formatted_lines)
        
    def print_route_sheets(self):
        try:
            self.status_label.configure(
                text="⏳ Printing route sheets...",
                text_color="#2563EB"
            )
            self.root.update()
            
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            
            for file in self.generated_files:
                workbook = excel.Workbooks.Open(file)
                sheet = workbook.Sheets(1)
                
                sheet.PageSetup.PrintArea = "A1:K44"
                sheet.PageSetup.Zoom = False
                sheet.PageSetup.FitToPagesWide = 1
                sheet.PageSetup.FitToPagesTall = 1
                
                workbook.PrintOut()
                workbook.Close(SaveChanges=False)
                
            excel.Quit()
            
            self.status_label.configure(
                text="✅ All route sheets printed successfully",
                text_color="#059669"
            )
            
        except Exception as e:
            self.status_label.configure(
                text="❌ Print error",
                text_color="#DC2626"
            )
            messagebox.showerror("Error", f"Failed to print route sheets: {str(e)}")
            
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = ModernRouteSheetApp()
    app.run()