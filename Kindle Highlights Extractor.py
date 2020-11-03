#pip install python-docx
import docx
import os
#os.rename('',r'file path\NEW file name.file type')
fo = open("My Clippings.txt",'r',encoding="utf8")
mydoc = docx.Document()
mydoc.add_heading("Highlights from my Books on Kindle", 0)
mydoc.save("Highlights from Kindle.docx")

L = fo.readlines()
#To remove the newline list items
L = list(filter(lambda a: a != '\n', L))

D = {}
for i in range(len(L)):
    try:
        if L[i] == '==========\n' and L[i+3] != "==========\n":            
            if str(L[i+1]) not in D:
                D[str(L[i+1])] = [L[i+3]]
            else:
                D[str(L[i+1])]+= [L[i+3]]
    except:
        #Exception raised when End of File is reached. So asking user to wait until all items of the file are saved.
        print('Please Wait')

for i in D:
    tempL = []
    count = 0
    mydoc.add_heading(i, 1)
    mydoc.save("Highlights from Kindle.docx")
    for j in D[i]:
        if j not in tempL:
            count+=1
            tempL+=[j]
            mydoc.add_paragraph(str(count)+'. '+j)
            mydoc.save("Highlights from Kindle.docx")
    D[i] = tempL
print('Done')
