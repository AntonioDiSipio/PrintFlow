import os
import json

SAVE_PATH = "printflow_files.json"

def load_file_list():
    if os.path.exists(SAVE_PATH):
        with open(SAVE_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def save_file_list(file_list):
    with open(SAVE_PATH, "w", encoding="utf-8") as f:
        json.dump(file_list, f, indent=2)

def add_file(file_list, filepath):
    if filepath.lower().endswith((".pdf", ".doc", ".docx")) and filepath not in file_list:
        file_list.append(filepath)

def remove_file(file_list, index):
    if 0 <= index < len(file_list):
        file_list.pop(index)