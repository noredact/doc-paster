2025-01-24
Doctor Paster

- This program reads a .csv document and allows user to press 1-9 on their number pad to send an element from that CSV to the cursor.
- A display shows what will be sent when the user presses a number.
- Documents that contain more than 9 elements can be cycled through using zero and dot (.) from the number pad.
- User can adjust how much zero or dot increments the program using ctrl + zero/dot
- For large documents, user can 'jump to' an element, a preview is provided. 
- Previews are provided to peek ahead/behind the current position the program is reading. 
- A section is provided to show other CSV documents contained in the current parent folder
- - User can use Numpad + (+/-) to navigate the current directory
- There are some features for users who have become acustomed to the program
  - Ctrl+(+/-) adjust the window's transparency
  - Shift+(+/-) adjust the text's darkness
    - Ctrl+Shift+(+/-) adjust both of these at the same time
  - The window can be shrunk to only show what is loaded to the number pad with Ctrl+h
- To disable the program so the number pad can be used regularly, press the * (multiply) key

# Known Issues
- This uses the clipboard to send the elements, so pressing keys too fast may result in unintended behavior
  - This is due to the ctrl key being held down by the program
- The first time the program launches some defaults don't load properly
- Text scrolling (text that doesn't fit in box scrolls)
  - A checkbox is supposed to start/stop this but only works if program launches with scrolling enabled
- Adjusting jump to/increment parameters requires mouse position to be exactly on the number
- Font size section requires fixing to match how adjust increment/ jump to position controls work
