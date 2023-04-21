import comtypes.client
import os
from PyPDF2 import PdfMerger
from tqdm import tqdm
import docx
import random
import string

def textToPDF(text, count=1):
    inputFile = f"{count}"
    cwd = os.getcwd()
    inputFile = os.path.join(cwd, inputFile)
    document = docx.Document()
    document.add_paragraph(text)
    document.save(inputFile)
    outputFile = inputFile.replace('.docx', '.pdf')
    if not os.path.exists('random'):
        os.makedirs('random')
    outputFile = os.path.join('random', outputFile)
    outputFile = os.path.join(cwd, outputFile)
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = 1
    doc = word.Documents.Open(inputFile)
    doc.SaveAs(outputFile, FileFormat=17)
    doc.Close()
    word.Quit()
    os.remove(inputFile)
    return outputFile
    
if __name__ == '__main__':
    n = int(input('Enter the number of pdfs required: '))
    for i in tqdm(range(n),desc='Generating pdfs', unit='PDFs'):
        letters = string.ascii_letters
        randomText= ''.join(random.choice(letters) for i in range(100000))
        textToPDF(randomText, count = i+1)
