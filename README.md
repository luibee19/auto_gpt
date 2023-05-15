# GPT Automaton

This is a Python program that automates interactions with the GPT-3 language model. It allows you to process a set of questions and generate answers using the OpenAI API.

## Features

- Browse and select an Excel file containing the questions.
- Process each question and generate answers using the GPT-3 model.
- Write the generated answers back to the Excel file.
- Display the progress and status messages in a GUI interface.

## Installation

1. Clone the repository:

git clone https://github.com/luibee19/auto_gpt.git

2. Install the required dependencies:
pip install -r requirements.txt

## Usage
Goto https://platform.openai.com/account/api-keys to create an API key if you havent had one
Save your API-key to api_key.txt at the same location of the code file
Run the auto_gpt_w_gui.py script: python auto_gpt_w_gui.py
The GUI window will appear.
Click the "Browse" button to select an Excel file containing the questions.
Click the "Process File" button to start the processing.
The program will generate answers for each question and write them back to the Excel file.
The progress and status messages will be displayed in the listbox on the GUI window.

## Contributing
Contributions are welcome! If you find any issues or have suggestions for improvements, please feel free to open an issue or submit a pull request.
