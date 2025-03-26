# -*- coding: utf-8 -*-
"""
Created on Wed Nov 20 14:47:14 2024

@author: Bram
"""

# Imports
import os
import pandas as pd
import tkinter as tk
import re
from tkinterdnd2 import TkinterDnD, DND_FILES
from anonemail import Email

#Read excel file

# File name of the Excel file
file_name_drank = "Verbruik_tabel_drank.xlsx"
#file_name_food = "Verbruik_tabel_food.xlsx"

drank_data = None
food_data = None
usage_table_drink = pd.DataFrame()
usage_table_food = pd.DataFrame()

try:
    # Read the Excel file
    drank = pd.read_excel(file_name_drank)
    drank_data = drank.dropna()
    drank_data["Hoeveelheid per bestelunit"] = drank_data["Hoeveelheid per bestelunit"].astype(float)

    # Display the content of the Excel file
    # print("Excel file content:")
    # print(drank_data)

except FileNotFoundError:
    print(f"Error: The file '{file_name_drank}' was not found in the current folder.")
except Exception as e:
    print(f"An error occurred: {e}")
    
#try:
#    # Read the Excel file
#    food_data = pd.read_excel(file_name_food)
#
#    # Display the content of the Excel file
#    print("Excel file content:")
#    print(food_data)
#
#except FileNotFoundError:
#    print(f"Error: The file '{file_name_food}' was not found in the current folder.")
#except Exception as e:
#    print(f"An error occurred: {e}")
    

#Make GUI:

# Initialize a dictionary to store the file paths
file_paths = {"plane1": None, "plane2": None, "plane3": None}
week_usage = {"plane1": None, "plane2": None, "plane3": None}

def remove_html_tags(text):
    """Remove html tags from a string"""
    clean = re.compile('<.*?>')
    return re.sub(clean, '', text)

def make_array(content):
    content = remove_html_tags(content)
    content = content.split('\n')
    content = [item for item in content if item.strip()]
    
    return content

def fill_usage_table(content, plane_key):
    # if(drink):
    #     usage = #make_drink_usage_table(104, width)
    # else:
    #     usage = make_food_usage_table(37, width)
        
    # usage = drank_data.iloc[:, :1].copy()  # Make a copy to avoid modifying the original DataFrame
    usage_week = []
    
    
    for i in range(len(drank_data)):
        # print('i: ', i)
        matched = False
        item = drank_data.iloc[i]['Product']
        # print(item + "!!!!")
        if not matched:
            for line in content:
                # print(line[:len(item)+3])
                # print(])
                if item in line[:len(item)+3]:
                    # print("Check")
                    # print(item)
                    # print(line[:len(item)])
                    # print(line)
                    split_line = line.split()
                    for elem in split_line:
                        if elem.isnumeric():
                            usage_week.append(float(elem))
                            matched = True
                            break
            if not matched:
                usage_week.append(0)  # Append 0 if no match is found
        else:
            break
        
    # print(len(usage_week))
    
    return usage_week

def week_usage_table(drink, file_name, plane_key):
    file = open(file_name, "r")
    content = file.read()
    file.close()
    content = make_array(content)
    table = fill_usage_table(content, plane_key)
    # print(plane_key[-1])
    # for i in range(len(table)):
    #     if(drink):
    #         # print(table[i])
    #         # print(drank_data.iloc[i]["Hoeveelheid per bestelunit"])
    #         table[i] = table[i]/drank_data.iloc[i]["Hoeveelheid per bestelunit"]
            
    return table

def process_file(event, label, plane_key):
    """Handle file drops and update the label and file_paths."""
    file_path = event.data.strip()
    if file_path.endswith(".htm"):
        file_paths[plane_key] = file_path  # Save the file path
        week_usage[plane_key] = week_usage_table(True, file_path, plane_key)
        label.config(text=f"Loaded: {file_path}")
    else:
        label.config(text="Error: Not an HTML file!")

def reset_planes():
    """Reset all planes to their default state."""
    plane1_label.config(text="Drop HTML Week 1")
    plane2_label.config(text="Drop HTML Week 2")
    plane3_label.config(text="Drop HTML Week 3")
    # Clear stored file paths
    for key in file_paths:
        file_paths[key] = None
        
# Function to send the Excel file via email
def send_email_with_attachment(dataframe, recipient_email, subject, body, attachment_filename):
    # Save DataFrame to Excel
    save_dataframe_to_excel(dataframe, attachment_filename)
    
    # Send email with attachment
    try:
        email = Email(
            sender_name="Anonymous Sender",
            sender_email="anon@example.com",  # Placeholder
            receiver_email=recipient_email,
            subject=subject,
            message=body,
            attachment=attachment_filename
        )
        email.send()
        print(f"Email sent to {recipient_email} with attachment: {attachment_filename}")
    except Exception as e:
        print(f"Error sending email: {e}")
    finally:
        # Clean up by removing the temporary file
        if os.path.exists(attachment_filename):
            os.remove(attachment_filename)
            
# Function to save DataFrame as an Excel file
def save_dataframe_to_excel(dataframe, filename):
    try:
        dataframe.to_excel(filename, index=False)
        print(f"DataFrame saved to {filename}")
    except Exception as e:
        print(f"Error saving DataFrame to Excel: {e}")


def process_files():
    """Print the saved file paths (for demonstration)."""
    usage_table_drink["Product"] = drank_data["Product"]
    three_weeks = []
    average = []
    maximum = []
    minimum = []
    for i in range(len(usage_table_drink)):
        three_weeks.append((week_usage["plane1"][i]+week_usage["plane2"][i]+week_usage["plane3"][i]))
        average.append(three_weeks[i]/3.0/drank_data.iloc[i]["Hoeveelheid per bestelunit"])
        maximum.append(max(week_usage["plane1"][i],week_usage["plane2"][i], week_usage["plane3"][i])/drank_data.iloc[i]["Hoeveelheid per bestelunit"])
        minimum.append(min(week_usage["plane1"][i],week_usage["plane2"][i], week_usage["plane3"][i])/drank_data.iloc[i]["Hoeveelheid per bestelunit"])
    print(len(three_weeks))
    print(len(average))
    print(len(minimum))
    print(len(maximum))
    usage_table_drink["Max per week"]=maximum
    usage_table_drink["Gemiddelde Bestelunit"]=average
    usage_table_drink["Min per week"]=minimum
    usage_table_drink["Verbruik 3 weken"]=three_weeks
    for key, week in week_usage.items():
        usage_table_drink.insert(int(key[-1])+4, f'Week {key[-1]}', week)
    # usage_table_drink.insert(1, "Verbruik 3 weken", three_weeks)
    
    print(usage_table_drink.to_markdown())
    
    # Save and send the Excel file
    recipient = "inkoop@cafedeklep.nl"
        
    subject = "Usage Table Drink"
        
    body = "Attached is the Usage Table Drink for the past 3 weeks."
    attachment_filename = "Usage_Table_Drink.xlsx"
    send_email_with_attachment(usage_table_drink, recipient, subject, body, attachment_filename)
        
        

# Initialize the TkinterDnD application
root = TkinterDnD.Tk()

# Set the window title and size
root.title("Drag and Drop HTML Files")
root.geometry("600x500")

# Create a frame to hold the three panes and the buttons
frame = tk.Frame(root)
frame.pack(expand=True, fill=tk.BOTH, pady=10)

# Define the labels for each plane
plane1_label = tk.Label(frame, text="Drop HTML Week 1", bg="lightblue", relief=tk.RAISED, height=5)
plane2_label = tk.Label(frame, text="Drop HTML Week 2", bg="lightgreen", relief=tk.RAISED, height=5)
plane3_label = tk.Label(frame, text="Drop HTML Week 3", bg="lightpink", relief=tk.RAISED, height=5)

# Pack the planes vertically
plane1_label.pack(fill=tk.BOTH, expand=True, pady=5)
plane2_label.pack(fill=tk.BOTH, expand=True, pady=5)
plane3_label.pack(fill=tk.BOTH, expand=True, pady=5)

# Enable drag-and-drop on each label
plane1_label.drop_target_register(DND_FILES)
plane2_label.drop_target_register(DND_FILES)
plane3_label.drop_target_register(DND_FILES)

# Bind drag-and-drop events to their respective planes
plane1_label.dnd_bind("<<Drop>>", lambda event: process_file(event, plane1_label, "plane1"))
plane2_label.dnd_bind("<<Drop>>", lambda event: process_file(event, plane2_label, "plane2"))
plane3_label.dnd_bind("<<Drop>>", lambda event: process_file(event, plane3_label, "plane3"))

# Create a frame for the buttons
button_frame = tk.Frame(root)
button_frame.pack(fill=tk.BOTH, pady=10)

# Add Reset and Print Paths buttons
reset_button = tk.Button(button_frame, text="Reset", command=reset_planes, bg="lightgrey")
print_button = tk.Button(button_frame, text="Process files", command=process_files, bg="lightgrey")

reset_button.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=5)
print_button.pack(side=tk.RIGHT, expand=True, fill=tk.BOTH, padx=5)

# Run the Tkinter event loop
root.mainloop()
