import os
from docx import Document

# Allow for the replacer to take in a dictionary of the field_to_replace (key) and the replacement_text (value)

def replacer(document, fields: dict[str, str]):
    for paragraph in document.paragraphs:
        for field_key in fields.keys():
            if field_key in paragraph.text:
                paragraph.text = paragraph.text.replace(field_key, fields.get(field_key))

    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for field_key in fields.keys():
                    if field_key in cell.text:
                        cell.text = cell.text.replace(field_key, fields.get(field_key))

def find_and_replace_field_single_file(doc_path, modified_path_ques):
    document = Document(doc_path.strip())

    fields_dict: dict[str, str] = dict()

    while True:
        field_to_replace = input("Enter the field to replace (enter Q to quit): ")

        if field_to_replace == 'Q' or field_to_replace == 'q':
            break

        replacement_text = input(f"Enter what to change {field_to_replace} to: ")
        fields_dict[field_to_replace] = replacement_text

        replacer(document, fields_dict)

    if len(fields_dict) > 0:
        print("\nField changes:")
        for field, replacement in fields_dict.items():
            print(f"{field} -> {replacement}")

        confirm = input("\nDo you want to apply these changes? (Y/N): ")
        if confirm.upper() == 'Y':
            if modified_path_ques == "Y":
                doc_dir, doc_filename = os.path.split(doc_path)
                modified_doc_path = os.path.join(doc_dir + "/modified-" + doc_filename)
                document.save(modified_doc_path)
                print("Field replacement complete. Modified document saved as:", modified_doc_path)
            elif modified_path_ques == "N":
                doc_path = doc_path.strip()
                document.save(doc_path)
                print("Field replacement complete. Modified document saved as:", doc_path)
        else:
            find_and_replace_field_single_file(doc_path, modified_path_ques)
    else:
        print("No changes applied. Exiting program.")

def find_and_replace_field_folder(folder_directory, fields: dict[str, str]):
    for root, _, files in os.walk(folder_directory):
        for file in files:
            file_path = os.path.join(root, file)
            if file_path.endswith('.docx'):
                doc = Document(file_path)
                replacer(doc, fields)
                doc.save(file_path)
                print("Field replacement complete. Modified document saved as:", file_path)

def apply_replacements(fields_dict, doc_path):
    if len(fields_dict) > 0:
        print("\nField changes:")
        for field, replacement in fields_dict.items():
            print(f"{field} -> {replacement}")

        confirm = input("\nDo you want to apply these changes? (Y/N): ")
        if confirm.upper() == 'Y':
            for root, _, files in os.walk(doc_path):
                for file in files:
                    file_path = os.path.join(root, file)
                    if file_path.endswith('.docx'):
                        doc = Document(file_path)
                        replacer(doc, fields_dict)
                        doc.save(file_path)
                        print("Field replacement complete. Modified document saved as:", file_path)
        elif confirm.upper() == 'N':
            for field, replacement in fields_dict.items():
                replacement_text = input(f"Enter what to change {replacement} to: ")
                fields_dict[field] = replacement_text
            apply_replacements(fields_dict, doc_path)
        else:
            print("No changes applied. Exiting program.")
    else:
        print("No changes applied. Exiting program.")

def replacements(doc_path):
    fields_dict: dict[str, str] = dict()

    while True:
        field_to_replace = input("Enter the field to replace (enter Q to quit): ")

        if field_to_replace == 'Q' or field_to_replace == 'q':
            break

        replacement_text = input(f"Enter what to change {field_to_replace} to: ")

        fields_dict[field_to_replace] = replacement_text

    apply_replacements(fields_dict, doc_path)

def main():
    file_or_folder_quest = input("Will this be a folder (Y) or file (N) (Y/N): ").upper()
    doc_path = input("Enter the path to the Word document: ")
    if file_or_folder_quest == "N":
        modified_path_ques = input("Do you want a new modified file (Y/N): ")
        find_and_replace_field_single_file(doc_path, modified_path_ques)
    elif file_or_folder_quest == "Y":
        replacements(doc_path)

if __name__ == "__main__":
    main()
