import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from docx import Document
import json
import os


class TemplateApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Template Application")
        self.templates = self.load_templates()

        # Template Selection Section
        self.template_var = tk.StringVar()
        self.template_label = tk.Label(root, text="Choose Template:")
        self.template_label.pack()
        self.template_dropdown = ttk.Combobox(root, textvariable=self.template_var)
        self.template_dropdown['values'] = list(self.templates.keys())
        self.template_dropdown.pack()
        self.template_dropdown.bind("<<ComboboxSelected>>", self.load_fields)

        # Fields Section
        self.fields_frame = tk.Frame(root)
        self.fields_frame.pack()

        # Save Button
        self.save_button = tk.Button(root, text="Save Document", command=self.save_document)
        self.save_button.pack()

    def load_templates(self):
        with open('D:\\Gayana Mahendra\\TemplateApplication\\templates.json', 'r') as f:
            return json.load(f)

    def load_fields(self, event):
        for widget in self.fields_frame.winfo_children():
            widget.destroy()

        self.fields = {}
        template = self.template_var.get()
        if template:
            for field in self.templates[template]["fields"]:
                label = tk.Label(self.fields_frame, text=field.capitalize() + ":")
                label.pack()
                entry = tk.Entry(self.fields_frame)
                entry.pack()
                self.fields[field] = entry

    def save_document(self):
        template_name = self.template_var.get()
        if not template_name:
            messagebox.showerror("Error", "Please select a template")
            return

        template_info = self.templates[template_name]
        doc = Document(template_info["file"])

        for field, entry in self.fields.items():
            value = entry.get()
            for paragraph in doc.paragraphs:
                if f'{{{{ {field} }}}}' in paragraph.text:
                    paragraph.text = paragraph.text.replace(f'{{{{ {field} }}}}', value)

        save_path = filedialog.asksaveasfilename(defaultextension=".docx",
                                                 filetypes=[("Word files", "*.docx")],
                                                 initialfile=f"{self.fields['name'].get()}_{template_name}")
        if save_path:
            doc.save(save_path)
            messagebox.showinfo("Success", "Document saved successfully!")


if __name__ == "__main__":
    root = tk.Tk()
    app = TemplateApp(root)
    root.mainloop()
