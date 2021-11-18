import os
import comtypes.client
import time


format_code = 17 

start_time = time.time()


# create the MS Word App
word_app = comtypes.client.CreateObject("Word.Application")

# Conversion from MS Word to pdf
file_input = os.path.abspath("./proposal_thesis.docx")
file_output = os.path.abspath("./proposal_thesis2.pdf")

word_file = word_app.Documents.Open(file_input)
word_file.SaveAs(file_output, FileFormat = format_code) 

# Close file and application 
word_file.close()
word_app.Quit()

end_time = time.time()  

print(f"total time for conversion = {end_time - start_time}")
