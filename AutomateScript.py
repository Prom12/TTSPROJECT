import threading

from Constants import TTS_CONTENT_PATH,TTS_RECORDERS_PATH,SIDE_BY_SIDE,GET_MOUSE_LOCATION
from WordScript import control_word, open_word,control_word_side
from SharedGlobalVariables import Singleton
from SharedFunctions import position_windows_side_by_side
from ExcelScript import read_excel_and_display,read_recorder_and_display 
from AdobeAuditionScript import open_adobe_audition



def main():
    # Get mouse location if needed
    GET_MOUSE_LOCATION()

    # Set the username
    Singleton.username = input('Who is recording ? ').lower()

    # Read data from Excel files
    TTS_Content = read_excel_and_display(TTS_CONTENT_PATH)
    TTS_Recorders = read_recorder_and_display(TTS_RECORDERS_PATH)

    # Start threads for opening Adobe Audition and Word
    adobe_audition_thread = threading.Thread(target=open_adobe_audition)
    word_thread = threading.Thread(target=open_word)
    adobe_audition_thread.start()
    word_thread.start()

    # Wait for both threads to finish opening the apps
    adobe_audition_thread.join()
    word_thread.join()

    # Position Word and Adobe Audition windows side by side
    # SIDE_BY_SIDE(position_windows_side_by_side())

    # Check for existing statements in TTS_Recorders
    existing_statements = set((row[2], row[3]) for row in TTS_Recorders[1:])

    # Process each row in TTS_Content
    for row in TTS_Content[1:]:
        statement_id, statement = row[0], row[1]
        statement_tuple = (statement, statement_id)

        # If the statement is not in TTS_Recorders, add it and control Word and Adobe Audition
        if statement_tuple not in existing_statements:
            TTS_Recorders.append(row)
            control_word_side(statement, statement_id)
            # control_word(statement, statement_id)
            existing_statements.add(statement_tuple)


if __name__ == "__main__":
    main()