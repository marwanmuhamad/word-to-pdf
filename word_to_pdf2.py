import os
import win32com.client
import time
from datetime import datetime

format_code = 17 

start_time = time.time()

today = datetime.today()
today = today.strftime("%d-%m-%Y")  
# create the MS Word App
word_app = win32com.client.Dispatch("Word.Application")

# Conversion from MS Word to pdf
file_input = os.path.abspath("./proposal_thesis.docx")
file_output = os.path.abspath(f"./proposal_thesis3-{today}.pdf")

word_file = word_app.Documents.Open(file_input)
word_file.SaveAs(file_output, FileFormat = format_code) 

# Close file and application 
word_file.Close()
word_app.Quit()

end_time = time.time()  

print(f"total time for conversion = {end_time - start_time}")
print(today)
