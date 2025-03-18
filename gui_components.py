from PyQt5.QtWidgets import (QLabel, QLineEdit, QPushButton, QTextEdit,
                           QVBoxLayout, QHBoxLayout, QFrame, QWidget,
                           QScrollArea, QMessageBox, QFileDialog)  # Import missing modules
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPalette, QColor, QFont, QPixmap, QIcon  # Import missing modules
from datetime import datetime
import os
import sys

def apply_styles(widget):
    widget.setStyleSheet("")  # Clear any inline styles
    widget.setProperty("class", widget.__class__.__name__)  # Add a class based on the widget type

def create_styled_label(text, is_header=False):
    label = QLabel(text)
    if is_header:
        label.setProperty("class", "header")
    return label

def create_styled_button(text):
    button = QPushButton(text)
    return button

def create_styled_input(placeholder=""):
    input_field = QLineEdit()
    input_field.setPlaceholderText(placeholder)
    return input_field

def create_styled_text_edit(placeholder=""):
    text_edit = QTextEdit()
    text_edit.setPlaceholderText(placeholder)
    return text_edit

def create_section_frame():
    frame = QFrame()
    frame.setProperty("class", "section")
    return frame

def create_api_section(parent):
    frame = create_section_frame()
    layout = QVBoxLayout(frame)
    
    api_label = create_styled_label("Gemini API Settings", True)
    layout.addWidget(api_label)
    
    input_frame = QFrame()
    input_layout = QHBoxLayout(input_frame)
    
    api_key_label = create_styled_label("API Key:")
    input_layout.addWidget(api_key_label)
    
    api_key_entry = create_styled_input("Enter your API key")
    api_key_entry.setEchoMode(QLineEdit.Password)
    input_layout.addWidget(api_key_entry)
    
    save_api_button = create_styled_button("Save API Key")
    input_layout.addWidget(save_api_button)
    
    layout.addWidget(input_frame)
    
    return frame, api_key_entry, save_api_button  # Return only what's needed

def create_event_details_section():
    frame = create_section_frame()
    layout = QVBoxLayout(frame)
    
    event_label = create_styled_label("Event Details", True)
    layout.addWidget(event_label)

    title_label = create_styled_label("Event Title:")
    title_entry = create_styled_input("Enter event title")
    
    date_label = create_styled_label("Event Date:")
    date_entry = create_styled_input(datetime.now().strftime('%Y-%m-%d'))
    
    time_label = create_styled_label("Event Time:")
    time_entry = create_styled_input(datetime.now().strftime('%H:%M'))
    
    description_label = create_styled_label("Event Description:")
    description_text = create_styled_text_edit("Enter event description")
    
    venue_label = create_styled_label("Venue:")
    venue_entry = create_styled_input("Enter venue")
    
    club_label = create_styled_label("Name of Student Led Club:")
    club_entry = create_styled_input("Enter club name")

    # Add all widgets to layout
    for widget in [title_label, title_entry, date_label, date_entry,
                  time_label, time_entry, description_label, description_text,
                  venue_label, venue_entry, club_label, club_entry]:
        layout.addWidget(widget)

    return (frame, title_entry, date_entry, time_entry, description_text,
            venue_entry, club_entry)  # Return the frame

def create_image_sections():
    frame = create_section_frame()
    layout = QVBoxLayout(frame)
    
    flyer_label = create_styled_label("Event Flyer", True)
    add_flyer_button = create_styled_button("Add Event Flyer")
    
    images_label = create_styled_label("Event Pictures", True)
    add_images_button = create_styled_button("Add Images")
    
    image_preview_frame = QFrame()
    image_preview_layout = QHBoxLayout(image_preview_frame)
    
    # Add widgets to layout
    layout.addWidget(flyer_label)
    layout.addWidget(add_flyer_button)
    layout.addWidget(images_label)
    layout.addWidget(add_images_button)
    layout.addWidget(image_preview_frame)

    return (frame, add_flyer_button, add_images_button, image_preview_frame, image_preview_layout)  # Return the frame

def create_attendance_section():
    frame = create_section_frame()
    layout = QVBoxLayout(frame)
    
    attendance_label = create_styled_label("Participants List", True)
    upload_attendance_button = create_styled_button("Upload Participants List (Excel/CSV)")
    
    layout.addWidget(attendance_label)
    layout.addWidget(upload_attendance_button)

    return frame, upload_attendance_button  # Return the frame

def create_preview_section():
    frame = create_section_frame()
    layout = QVBoxLayout(frame)
    
    preview_label = create_styled_label("Live Preview", True)
    preview_edit = create_styled_text_edit()
    preview_edit.setReadOnly(True)
    
    layout.addWidget(preview_label)
    layout.addWidget(preview_edit)

    return frame, preview_edit  # Return the frame
