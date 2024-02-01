"""
Created to call environ email variable to log in to snowflake

"""

from tkinter import Tk,Text,Button
import variables
import os

class GetEmail(Tk):
    def __init__(self):
        variables.set_email_file()

        super().__init__()
        self.title("Save your e-mail to log in on SnowFlake")
        self.textBox = Text(self, height=1, width=50)
        self.textBox.pack()
        self.buttonCommit = Button(self, height=1, width=10, text="Save E-mail!",
                              command=lambda: self.retrieve_input())
        # command=lambda: retrieve_input() >>> just means do this when i press the button
        self.buttonCommit.pack()


    def retrieve_input(self):
        inputValue = self.textBox.get("1.0", "end-1c")
        with open(f'{os.environ["USERPROFILE"]}\\.SpartaTool\\sparta_email.txt', 'w') as f:
            f.write(inputValue)
            f.close()

        self.quit()
