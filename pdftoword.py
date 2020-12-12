import win32com.client, os
print('''Tips:
1. You should copy the converter file into same Directory.
2. If you don't follow tips 1 then try to enter file name with file location.
Thank you!
\n\n\n
Created By Hamid Ul Bari.
''')
word=win32com.client.Dispatch("word.Application")
word.visible=0
doc_pdf=input("Enter the file Name: ")
input_file=os.path.abspath(doc_pdf)
wb=word.Documents.Open(input_file)
output_file=os.path.abspath(doc_pdf[0:-4]+"docx".format())
wb.SaveAs2(output_file,FileFormat=16)
print("Converted Succesfully!")
wb.Close()
word.Quit()
