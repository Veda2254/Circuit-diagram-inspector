"""
Circuit Diagram Inspector
Interactive PDF inspection tool for circuit diagrams with error logging to Excel
"""

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, Menu
from PIL import Image, ImageTk, ImageDraw
import fitz  # PyMuPDF
from openpyxl import Workbook, load_workbook
from datetime import datetime
import os
import json


class CircuitInspector:
    def __init__(self, root):
        self.root = root
        self.root.title("Circuit Diagram Inspector")
        self.root.geometry("1400x900")
        
        # Data storage
        self.pdf_document = None
        self.current_page = 0
        self.cabinet_id = ""
        self.annotations = []  # List of {type: 'ok'/'error', x, y, error_text}
        self.excel_file = "inspection_log.xlsx"
        self.zoom_level = 1.0
        
        # Error categories
        self.error_categories = {
            'Wire': ['Wire wrong', 'Ferrule direction wrong', 'Wiring not present'],
            'Fuse': ['Fuse missing', 'Wrong fuse rating', 'Fuse orientation wrong'],
            'Component': ['Missing component', 'Wrong material installation', 'Missing material', 'Wrong component type'],
            'General': ['Assembly error', 'Labeling error', 'Connection loose', 'Other']
        }
        
        self.setup_ui()
        self.setup_excel()
        
    def setup_ui(self):
        """Create the user interface"""
        # Top toolbar
        toolbar = tk.Frame(self.root, bg='#2c3e50', height=60)
        toolbar.pack(side=tk.TOP, fill=tk.X)
        
        # Buttons
        btn_style = {'bg': '#3498db', 'fg': 'white', 'padx': 15, 'pady': 8, 'font': ('Arial', 10)}
        
        tk.Button(toolbar, text="📁 Load PDF", command=self.load_pdf, **btn_style).pack(side=tk.LEFT, padx=5, pady=10)
        tk.Button(toolbar, text="🆔 Set Cabinet ID", command=self.set_cabinet_id, **btn_style).pack(side=tk.LEFT, padx=5, pady=10)
        
        # Cabinet ID display
        self.cabinet_label = tk.Label(toolbar, text="Cabinet: Not Set", bg='#2c3e50', fg='white', font=('Arial', 11, 'bold'))
        self.cabinet_label.pack(side=tk.LEFT, padx=20)
        
        # Page navigation
        tk.Button(toolbar, text="◀ Prev", command=self.prev_page, bg='#95a5a6', fg='white', padx=10, pady=8).pack(side=tk.LEFT, padx=5)
        self.page_label = tk.Label(toolbar, text="Page: 0/0", bg='#2c3e50', fg='white', font=('Arial', 10))
        self.page_label.pack(side=tk.LEFT, padx=5)
        tk.Button(toolbar, text="Next ▶", command=self.next_page, bg='#95a5a6', fg='white', padx=10, pady=8).pack(side=tk.LEFT, padx=5)
        
        # Zoom controls
        tk.Button(toolbar, text="🔍+", command=self.zoom_in, bg='#27ae60', fg='white', padx=10, pady=8).pack(side=tk.LEFT, padx=(20, 2))
        tk.Button(toolbar, text="🔍-", command=self.zoom_out, bg='#27ae60', fg='white', padx=10, pady=8).pack(side=tk.LEFT, padx=2)
        
        # Save/Export buttons
        tk.Button(toolbar, text="💾 Save Annotations", command=self.save_annotations, bg='#e67e22', fg='white', padx=10, pady=8).pack(side=tk.RIGHT, padx=5, pady=10)
        tk.Button(toolbar, text="📊 Open Excel", command=self.open_excel, bg='#16a085', fg='white', padx=10, pady=8).pack(side=tk.RIGHT, padx=5, pady=10)
        
        # Main canvas with scrollbars
        canvas_frame = tk.Frame(self.root)
        canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbars
        v_scrollbar = tk.Scrollbar(canvas_frame, orient=tk.VERTICAL)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        h_scrollbar = tk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Canvas
        self.canvas = tk.Canvas(canvas_frame, bg='#ecf0f1', 
                               yscrollcommand=v_scrollbar.set,
                               xscrollcommand=h_scrollbar.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        v_scrollbar.config(command=self.canvas.yview)
        h_scrollbar.config(command=self.canvas.xview)
        
        # Bind mouse events
        self.canvas.bind("<Button-1>", self.on_left_click)
        self.canvas.bind("<Button-3>", self.on_right_click)
        
        # Instructions panel
        instructions = tk.Frame(self.root, bg='#34495e', height=50)
        instructions.pack(side=tk.BOTTOM, fill=tk.X)
        
        inst_text = "🖱️ Left Click: Mark as OK (Green) | Right Click: Report Error (Yellow) | Use toolbar to navigate"
        tk.Label(instructions, text=inst_text, bg='#34495e', fg='white', 
                font=('Arial', 10), pady=15).pack()
    
    def setup_excel(self):
        """Initialize or load Excel file"""
        if not os.path.exists(self.excel_file):
            wb = Workbook()
            ws = wb.active
            ws.title = "Inspection Log"
            ws.append(["Timestamp", "Cabinet ID", "Page", "Component Type", "Error Description (Component + Error)", "Inspector"])
            wb.save(self.excel_file)
    
    def load_pdf(self):
        """Load a PDF file"""
        file_path = filedialog.askopenfilename(
            title="Select Circuit Diagram PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                self.pdf_document = fitz.open(file_path)
                self.current_page = 0
                self.annotations = []
                self.zoom_level = 1.0
                
                # Try to extract cabinet ID from filename
                filename = os.path.basename(file_path)
                suggested_id = filename.replace('.pdf', '').replace('_', '-')
                self.cabinet_id = suggested_id
                self.cabinet_label.config(text=f"Cabinet: {self.cabinet_id}")
                
                self.display_page()
                messagebox.showinfo("Success", f"Loaded PDF with {len(self.pdf_document)} pages")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load PDF: {str(e)}")
    
    def set_cabinet_id(self):
        """Set the cabinet identifier"""
        new_id = simpledialog.askstring("Cabinet ID", "Enter Cabinet ID:", 
                                       initialvalue=self.cabinet_id)
        if new_id:
            self.cabinet_id = new_id
            self.cabinet_label.config(text=f"Cabinet: {self.cabinet_id}")
    
    def display_page(self):
        """Render and display the current PDF page"""
        if not self.pdf_document:
            return
        
        try:
            # Get page
            page = self.pdf_document[self.current_page]
            
            # Render at higher resolution for better quality
            mat = fitz.Matrix(2.0 * self.zoom_level, 2.0 * self.zoom_level)
            pix = page.get_pixmap(matrix=mat)
            
            # Convert to PIL Image
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            # Draw annotations
            draw = ImageDraw.Draw(img)
            for ann in self.annotations:
                if ann['page'] == self.current_page:
                    x, y = int(ann['x'] * self.zoom_level * 2), int(ann['y'] * self.zoom_level * 2)
                    size = int(15 * self.zoom_level)
                    
                    if ann['type'] == 'ok':
                        # Green checkmark
                        color = 'green'
                        draw.ellipse([x-size, y-size, x+size, y+size], outline=color, width=3, fill=(0, 255, 0, 100))
                        draw.text((x-size//2, y-size//2), "✓", fill='darkgreen')
                    else:
                        # Yellow error marker
                        color = 'yellow'
                        draw.ellipse([x-size, y-size, x+size, y+size], outline='orange', width=3, fill=(255, 255, 0, 150))
                        draw.text((x-size//2, y-size//2), "!", fill='red')
            
            # Convert to PhotoImage
            self.photo = ImageTk.PhotoImage(img)
            
            # Update canvas
            self.canvas.delete("all")
            self.canvas.create_image(0, 0, anchor=tk.NW, image=self.photo)
            self.canvas.config(scrollregion=self.canvas.bbox(tk.ALL))
            
            # Update page label
            self.page_label.config(text=f"Page: {self.current_page + 1}/{len(self.pdf_document)}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to display page: {str(e)}")
    
    def on_left_click(self, event):
        """Handle left click - mark as OK"""
        if not self.pdf_document or not self.cabinet_id:
            messagebox.showwarning("Warning", "Please load a PDF and set Cabinet ID first")
            return
        
        # Get canvas coordinates
        x = self.canvas.canvasx(event.x) / (2.0 * self.zoom_level)
        y = self.canvas.canvasy(event.y) / (2.0 * self.zoom_level)
        
        # Add OK annotation
        self.annotations.append({
            'type': 'ok',
            'page': self.current_page,
            'x': x,
            'y': y,
            'timestamp': datetime.now().isoformat()
        })
        
        self.display_page()
    
    def on_right_click(self, event):
        """Handle right click - show error menu"""
        if not self.pdf_document or not self.cabinet_id:
            messagebox.showwarning("Warning", "Please load a PDF and set Cabinet ID first")
            return
        
        # Get canvas coordinates
        x = self.canvas.canvasx(event.x) / (2.0 * self.zoom_level)
        y = self.canvas.canvasy(event.y) / (2.0 * self.zoom_level)
        
        # Ask for component name/label
        component_name = simpledialog.askstring(
            "Component Name", 
            "Enter component name/label (e.g., 'F1 fuse', 'Wire X3-5'):",
            parent=self.root
        )
        
        if not component_name:
            return  # User cancelled
        
        # Create context menu
        menu = Menu(self.root, tearoff=0)
        
        for category, errors in self.error_categories.items():
            cat_menu = Menu(menu, tearoff=0)
            for error in errors:
                cat_menu.add_command(
                    label=error,
                    command=lambda e=error, c=category, cx=x, cy=y, cn=component_name: self.log_error(c, e, cx, cy, cn)
                )
            menu.add_cascade(label=f"🔧 {category}", menu=cat_menu)
        
        # Show menu at cursor position
        menu.tk_popup(event.x_root, event.y_root)
    
    def log_error(self, component_type, error_description, x, y, component_name):
        """Log an error to Excel and add annotation"""
        try:
            # Create detailed error message
            detailed_error = f"{component_name} {error_description}"
            
            # Add annotation
            self.annotations.append({
                'type': 'error',
                'page': self.current_page,
                'x': x,
                'y': y,
                'component': component_type,
                'component_name': component_name,
                'error': error_description,
                'timestamp': datetime.now().isoformat()
            })
            
            # Log to Excel
            wb = load_workbook(self.excel_file)
            ws = wb.active
            
            row_data = [
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                self.cabinet_id,
                self.current_page + 1,
                component_type,
                detailed_error,
                os.getlogin()
            ]
            
            ws.append(row_data)
            wb.save(self.excel_file)
            
            self.display_page()
            
            # Show confirmation
            self.root.after(100, lambda: messagebox.showinfo(
                "Logged", 
                f"Error logged: {self.cabinet_id} - {detailed_error}"
            ))
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to log error: {str(e)}")
    
    def prev_page(self):
        """Go to previous page"""
        if self.pdf_document and self.current_page > 0:
            self.current_page -= 1
            self.display_page()
    
    def next_page(self):
        """Go to next page"""
        if self.pdf_document and self.current_page < len(self.pdf_document) - 1:
            self.current_page += 1
            self.display_page()
    
    def zoom_in(self):
        """Increase zoom level"""
        if self.zoom_level < 3.0:
            self.zoom_level += 0.25
            self.display_page()
    
    def zoom_out(self):
        """Decrease zoom level"""
        if self.zoom_level > 0.5:
            self.zoom_level -= 0.25
            self.display_page()
    
    def save_annotations(self):
        """Save annotations to JSON file"""
        if not self.cabinet_id:
            messagebox.showwarning("Warning", "Please set Cabinet ID first")
            return
        
        try:
            save_file = f"{self.cabinet_id}_annotations.json"
            with open(save_file, 'w') as f:
                json.dump(self.annotations, f, indent=2)
            
            messagebox.showinfo("Saved", f"Annotations saved to {save_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save annotations: {str(e)}")
    
    def open_excel(self):
        """Open the Excel log file"""
        try:
            os.startfile(self.excel_file)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Excel: {str(e)}")


def main():
    root = tk.Tk()
    app = CircuitInspector(root)
    root.mainloop()


if __name__ == "__main__":
    main()
