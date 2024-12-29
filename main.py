import json
import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from docx import Document


class TemplateApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Template Application")

        # Center the application and set its size to 75% of the screen
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        app_width = int(screen_width * 0.75)
        app_height = int(screen_height * 0.75)

        x_pos = (screen_width - app_width) // 2
        y_pos = (screen_height - app_height) // 2

        self.root.geometry(f"{app_width}x{app_height}+{x_pos}+{y_pos}")
        self.root.resizable(True, True)

        self.templates = self.load_templates()

        # Determine the longest template name for width adjustment
        max_name_length = max(len(name) for name in self.templates.keys())

        # Template Selection Section
        self.template_var = tk.StringVar()
        self.template_label = tk.Label(root, text="Choose Template:")
        self.template_label.pack()

        self.template_dropdown = ttk.Combobox(root, textvariable=self.template_var, width=max_name_length + 5)
        self.template_dropdown['values'] = list(self.templates.keys())
        self.template_dropdown.pack()
        self.template_dropdown.bind("<<ComboboxSelected>>", self.load_fields)

        # Fields Section (Scrollable Frame)
        self.canvas = tk.Canvas(root)
        self.scrollbar = ttk.Scrollbar(root, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.scrollable_frame.bind("<Enter>", self._bind_mouse_scroll)
        self.scrollable_frame.bind("<Leave>", self._unbind_mouse_scroll)

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.bind("<Configure>", lambda e: self.recenter_scrollable_frame())

        self.scrollbar.pack(side="right", fill="y")

        self.fields_frame = self.scrollable_frame  # Use scrollable_frame as the parent for fields

        # Save Button (initially hidden)
        self.save_button = tk.Button(root, text="Save Document", command=self.save_document)

    def _bind_mouse_scroll(self, event):
        """Bind the mouse scroll event to the canvas when the cursor enters the scrollable frame."""
        self.canvas.bind_all("<MouseWheel>", self._on_mouse_scroll)

    def _unbind_mouse_scroll(self, event):
        """Unbind the mouse scroll event when the cursor leaves the scrollable frame."""
        self.canvas.unbind_all("<MouseWheel>")

    def _on_mouse_scroll(self, event):
        """Scroll the canvas vertically using the mouse wheel."""
        self.canvas.yview_scroll(-1 * (event.delta // 120), "units")

    def resource_path(self, relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)

    def load_templates(self):
        path = self.resource_path('templates.json')
        with open(path, 'r') as f:
            return json.load(f)

    def load_fields(self, event):
        # Clear the previous fields
        for widget in self.fields_frame.winfo_children():
            widget.destroy()

        self.fields = {}
        template = self.template_var.get()
        if template:
            row = 0  # Row counter for grid positioning

            # Get canvas width and calculate padding
            frame_width = self.scrollable_frame.winfo_width()
            self.fields_frame.configure(width=frame_width)

            # Center-align the fields and labels
            self.fields_frame.grid_columnconfigure(0, weight=1)  # Center labels
            self.fields_frame.grid_columnconfigure(1, weight=1)  # Center entry fields

            for field in self.templates[template]["fields"]:
                # Create label for the field
                label = tk.Label(self.fields_frame, text=field)
                label.grid(row=row, column=0, padx=10, pady=5, sticky="e")  # Align label to the right

                # Create entry for the field
                entry = tk.Entry(self.fields_frame)
                entry.grid(row=row, column=1, padx=10, pady=5, sticky="w")  # Align entry to the left

                self.fields[field] = entry
                row += 1  # Move to the next row

            # Add the Save button to the bottom and center it
            self.fields_frame.grid_rowconfigure(row, weight=1)  # Allow space for the button
            save_button_frame = tk.Frame(self.fields_frame)
            save_button_frame.grid(row=row, column=0, columnspan=2, pady=20)

            self.save_button = tk.Button(save_button_frame, text="Save Document", command=self.save_document)
            self.save_button.pack()

        # Reconfigure the canvas to include the padding
        self.recenter_scrollable_frame()

    def recenter_scrollable_frame(self):
        self.root.update_idletasks()  # Ensure dimensions are up-to-date

        # Calculate padding
        canvas_width = self.canvas.winfo_width()
        frame_width = self.scrollable_frame.winfo_width()
        x_offset = (canvas_width - frame_width) // 2

        # Update the position of the scrollable frame in the canvas
        self.canvas.delete("all")  # Clear existing canvas content
        self.canvas.create_window((x_offset, 0), window=self.scrollable_frame, anchor="nw")

    def save_document(self):
        template_name = self.template_var.get()
        if not template_name:
            messagebox.showerror("Error", "Please select a template")
            return

        template_info = self.templates[template_name]
        template_path = self.resource_path(template_info["file"])
        if not os.path.exists(template_path):
            messagebox.showerror("Error", f"Template file not found: {template_path}")
            return

        doc = Document(template_path)

        for field, entry in self.fields.items():
            value = entry.get()
            for paragraph in doc.paragraphs:
                if f'{{{{ {field} }}}}' in paragraph.text:
                    paragraph.text = paragraph.text.replace(f'{{{{ {field} }}}}', value)

        save_path = filedialog.asksaveasfilename(defaultextension=".docx",
                                                 filetypes=[("Word files", "*.docx")],
                                                 initialfile=f"{self.fields['Name'].get()}_{template_name}")
        if save_path:
            doc.save(save_path)
            messagebox.showinfo("Success", "Document saved successfully!")


if __name__ == "__main__":
    root = tk.Tk()
    app = TemplateApp(root)
    root.mainloop()
