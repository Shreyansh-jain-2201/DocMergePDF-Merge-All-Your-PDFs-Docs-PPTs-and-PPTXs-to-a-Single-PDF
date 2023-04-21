import comtypes.client
import os
from PyPDF2 import PdfMerger
from tqdm import tqdm


def DocToPdf(inputFile):
    outputFile = inputFile.replace('.docx', '.pdf')
    word = comtypes.client.CreateObject('Word.Application')
    cwd = os.getcwd()
    inputFile = os.path.join(cwd, inputFile)
    word.Visible = 1
    doc = word.Documents.Open(inputFile)
    doc.SaveAs(outputFile, FileFormat=17)
    doc.Close()
    word.Quit()
    return outputFile


def PPTtoPdf(inputFile):
    if inputFile.endswith('.ppt'):
        outputFile = inputFile.replace('.ppt', '.pdf')
    elif inputFile.endswith('.pptx'):
        outputFile = inputFile.replace('.pptx', '.pdf')
    powerpoint = comtypes.client.CreateObject('Powerpoint.Application')
    powerpoint.Visible = 1
    deck = powerpoint.Presentations.Open(inputFile)
    deck.SaveAs(outputFile, FileFormat=32)
    deck.Close()
    powerpoint.Quit()
    return outputFile


def MergePdf(inputFiles, outputFile):
    merger = PdfMerger()
    for pdf in tqdm(inputFiles, desc='Merging pdfs', unit='files'):
        merger.append(pdf)
    merger.write(outputFile)
    merger.close()
    return outputFile


def delete(filename):
    os.remove(filename)


if __name__ == '__main__':
    PDFs = []
    cwd = os.getcwd()
    files = [file for file in os.listdir('.') if file.endswith('.pdf') or file.endswith('.docx') or file.endswith('.pptx') or file.endswith('.ppt')]
    wantToRename = input('Do you want to rename the files? (y/n): ').lower()
    if (wantToRename == 'y'):
        for file in tqdm(files, desc='Renaming files', unit='files'):
            try:
                num = int(file.split('-')[0])
            except Exception:
                newFile = f'00{file}'
                files[files.index(file)] = newFile
                newFile = os.path.join(cwd, newFile)
                file = os.path.join(cwd, file)
                os.rename(file, newFile)
                file = newFile
                continue
            if (num < 10):
                newFile = f'0{file}'
                files[files.index(file)] = newFile
                newFile = os.path.join(cwd, newFile)
                file = os.path.join(cwd, file)
                os.rename(file, newFile)
                file = newFile
    for file in tqdm(files, desc='Converting files to pdf', unit='files'):
        file = os.path.join(cwd, file)
        if file.endswith('.docx'):
            try:
                pdf = DocToPdf(file)
                PDFs.append(pdf)
            except Exception:
                print('Error converting file from docx to pdf: ', file)
        elif file.endswith('.pptx'):
            try:
                pdf = PPTtoPdf(file)
                PDFs.append(pdf)
            except Exception:
                print('Error converting file from pptx to pdf: ', file)
        elif file.endswith('.ppt'):
            try:
                pdf = PPTtoPdf(file)
                PDFs.append(pdf)
            except Exception:
                print('Error converting file from ppt to pdf: ', file)

        elif file.endswith('.pdf'):
            PDFs.append(file)
    PDFs.sort()
    name = input('Enter the name of the output file: ')
    if not name.endswith('.pdf'):
        name += '.pdf'
    if not os.path.exists('output'):
        os.mkdir('output')
    name = os.path.join("output", name)
    name = os.path.join(cwd, name)
    MergePdf(PDFs, name)
    files += PDFs
    files = list(set(files))
    for file in tqdm(files, desc='Deleting files', unit='files'):
        if os.path.exists(file):
            delete(file)
