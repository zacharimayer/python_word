import win32com.client as win32
import csv
import os

def extract_track_changes(doc_path):
    # Open Microsoft Word via COM
    word_app = win32.Dispatch('Word.Application')
    word_app.Visible = False  # Keep the Word application hidden

    # Open the document
    doc = word_app.Documents.Open(doc_path)

    # Get all revisions (track changes)
    revisions = doc.Revisions

    changes_list = []
    for revision in revisions:
        change = {
            'Author': revision.Author,
            'Date': revision.Date,
            'Type': revision.Type,  # Type of the change (Insertion/Deletion/FormatChange etc.)
            'Text': revision.Range.Text  # The text that was changed
        }
        changes_list.append(change)

    # Close the document and quit Word
    doc.Close(False)
    word_app.Quit()

    return changes_list

def save_changes_to_csv(changes, csv_path):
    # Specify the headers for the CSV file
    headers = ['Author', 'Date', 'Type', 'Text']

    # Write changes to a CSV file
    with open(csv_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.DictWriter(file, fieldnames=headers)

        # Write the header
        writer.writeheader()

        # Write the changes row by row
        for change in changes:
            writer.writerow(change)

    print(f"Changes have been written to {os.path.abspath(csv_path)}")


if __name__ == "__main__":
    # Prompt the user to enter the path to the Word document
    doc_path = input("Please enter the full path of the Word document (.docx): ").strip()

    # Validate if the file exists
    if not os.path.isfile(doc_path) or not doc_path.lower().endswith('.docx'):
        print("Error: The specified file either does not exist or is not a .docx file.")
    else:
        # Extract the folder from the document path
        folder_path = os.path.dirname(doc_path)

        # Generate a CSV filename based on the document name
        csv_filename = os.path.splitext(os.path.basename(doc_path))[0] + '_track_changes.csv'

        # Create the full path for the CSV file
        csv_path = os.path.join(folder_path, csv_filename)

        # Extract track changes from the document
        changes = extract_track_changes(doc_path)

        # Save the changes to a CSV file in the same folder
        save_changes_to_csv(changes, csv_path)
