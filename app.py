import openpyxl
from tkinter import filedialog, Tk, Button, Label, Entry

def browse_excel():
  """
  Opens a file dialog box to select an excel file.
  """
  global excel_file_path
  excel_file_path = filedialog.askopenfilename(
      title="Select Excel File",
      filetypes=[("Excel Files", "*.xlsx *.xlsm *.xlsb")]
  )
  if excel_file_path:
    # Update label to show selected file path
    file_path_label.config(text=f"Selected File: {excel_file_path}")

def convert_to_pdf():
  """
  Converts the uploaded excel file to PDF if a path is provided.
  """
  if not excel_file_path:
    return  # No file selected, do nothing

  # Get output filename from entry field
  output_filename = output_filename_entry.get()

  try:
    # Load excel workbook
    workbook = openpyxl.load_workbook(excel_file_path)

    # Save entire workbook as PDF (each sheet becomes a separate page)
    workbook.save(f"{output_filename}.pdf")

    # Display success message
    message_label.config(text="Excel converted to PDF successfully!", fg="green")
  except FileNotFoundError:
    message_label.config(text="Error: Excel file not found.", fg="red")
  except Exception as e:
    message_label.config(text=f"Error: {str(e)}", fg="red")

# Initialize main window
root = Tk()
root.title("Excel to PDF Converter")

# Browse button
browse_button = Button(root, text="Browse Excel", command=browse_excel)
browse_button.pack(pady=10)

# File path label (initially empty)
file_path_label = Label(root, text="")
file_path_label.pack()

# Output filename entry
output_filename_label = Label(root, text="Output Filename:")
output_filename_label.pack()
output_filename_entry = Entry(root)
output_filename_entry.pack()

# Convert button
convert_button = Button(root, text="Convert to PDF", command=convert_to_pdf)
convert_button.pack(pady=10)

# Message label
message_label = Label(root, text="")
message_label.pack()

# Run the main loop
root.mainloop()

# Global variable to store the selected excel file path
excel_file_path = None
