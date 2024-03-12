## Python Software Data Extractor
Python script that scrapes software using the pyautogui library and stores the collected information in an xlsx file, everything is automated, just have the software open and run the script!

## :warning: Prerequisites

- [Pandas](https://pandas.pydata.org/docs/)

- [OS](https://docs.python.org/3/library/os.html)

- [Boto3](https://boto3.amazonaws.com/v1/documentation/api/latest/index.html)

- [Datetime](https://docs.python.org/3/library/datetime.html)
 
- [PySimpleGUI](https://www.pysimplegui.org/en/latest/)

- [Ctypes](https://docs.python.org/3/library/ctypes.html)

- [Pyperclip](https://pyperclip.readthedocs.io/en/latest/)

## WHAT THE SCRIPT DO:
- 1° Enters a AWS S3 bucket;
- 2° Takes all the names of the files inside the bucket;
- 3° Assembles a list and compares it with the current date;
- 4° If the current date is the same as the name of the file "horario" it shows online;
- 5° Else shows offline;
- 6° Shows the filtered list with online and offline asset management in a pysimplegui graphical interface;
- 7° Create an Excel file in the main folder.

## OBSERVATION

This code will only work if you redo the modifications of the dataframe that is taken from the cloud, it also needs to contain a folder called 'statusbolt' to save the excel file and an excel file inside that folder called 'idbolt' for it to replace the names with ids. 
Don't forget to put your aws credentials!

<p align="center">
    <img src="https://github.com/iagoapiai/Python-Software-Data-Extractor/assets/116030785/aaa92bd4-d7b9-42b3-9a83-bf3f91089ea9
">
</p>

Personal project, made to increase efficiency in my work! ❤️



