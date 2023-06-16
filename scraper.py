import numpy as np
import pandas as pd
import csv
from tkinter import *
from tkinter import filedialog


def getTsvFilepath():
    global tsvFileName
    tsvFileName = filedialog.askopenfilename(initialdir = "./",
                                          title = "Select a File")
    tsvLabel.configure(text="File Opened: "+tsvFileName)


def getMrpFilepath():
    global mrpFileName
    mrpFileName = filedialog.askopenfilename(initialdir = "./",
                                          title = "Select a File")
    mrpLabel.configure(text="File Opened: "+mrpFileName)
    
    

def getOutputFilepath():
    global outputFileName
    outputFileName = filedialog.asksaveasfilename(initialdir = "./",
                                          title = "Select a File",
                                          filetypes = (("Text files",
                                                        "*.txt*"),
                                                       ("all files",
                                                        "*.*")))
    outputLabel.configure(text="File Opened: "+outputFileName)


def processTsv():
    with open(tsvFileName) as fd:
        fileReader = csv.reader(fd,delimiter="\t")
        headers = list()
        for row in fileReader:
            if fileReader.line_num == 2:
                metaData = row
            elif fileReader.line_num == 5:
                headers.append(row[0:3])
            elif fileReader.line_num == 6:
                headers.append(row[3:15])
                flatHeaders = [j for sub in headers for j in sub]  # make headers 1D
                df = pd.DataFrame()
            elif fileReader.line_num >= 7:
                r = row[0:15]
                tempDf = pd.DataFrame(r)
                tempDf = tempDf.transpose()
                df = pd.concat([df,tempDf],ignore_index=True)

        df.columns = flatHeaders
        immiDF = pd.DataFrame()
        
        for index, row in df.iterrows():
            item = row['Item']
            buffer = int(row[flatHeaders[9]]) # existing buffer before months
            for header in flatHeaders[10:15]:
                quant = 0
                formattedHeader = header.replace("Month ","")
                quant = int(row[header])
                if quant != 0 and buffer != 0:
                    if quant >= buffer:           # first order with buffer
                        quant = quant - buffer
                        buffer = 0
                        continue
                    else:
                        quant = 0
                        buffer = buffer - quant

                if quant != 0:
                    entry = [item,quant,formattedHeader]
                    tempDF = pd.DataFrame(entry)
                    tempDF = tempDF.transpose()
                    immiDF = pd.concat([immiDF,tempDF],ignore_index=True)
                    
        immiDF.columns = ["Item", "Qty", "Date"]
    return immiDF


def processMrp():
    mrpSheet = pd.read_excel(mrpFileName)
    return mrpSheet


def createForecast():
    immiSheet = processTsv()
    mrpSheet = processMrp()
    with pd.ExcelWriter(outputFileName) as writer:
        immiSheet.to_excel(writer, sheet_name="IMMI", index=False)
        mrpSheet.to_excel(writer, sheet_name="Trimark", index=False)


def main():
    window = Tk()
    window.title('File Explorer')
    window.geometry("500x500")
    window.config(background = "white")

    global tsvLabel, mrpLabel,outputLabel
    tsvLabel = Label(window,
                    text = "Please select a .tsv file with IMMI Projections",
                    width = 50, height = 2,
                    fg = "blue", bg= "white")
    
    mrpLabel = Label(window,
                    text = "Please select an .xls file with MRP Projections",
                    width = 50, height = 2,
                    fg = "blue", bg="white")
    
    outputLabel = Label(window,
                    text = "Please select an .xlsx output file",
                    width = 50, height = 2,
                    fg = "blue", bg="white")
    
    tsvButton = Button(window,
                        text = "Browse for TSV file",
                        command = getTsvFilepath)
    
    mrpButton = Button(window,
                       text="Browse for MRP file",
                       command=getMrpFilepath)
    
    outputButton = Button(window,
                          text="Browse for Output file",
                          command=getOutputFilepath)
    
    createButton = Button(window,
                        text="Create Forecast",
                        command=createForecast)
    
    exitButton = Button(window,
                        text = "Exit",
                        command = exit)

    tsvLabel.grid(column = 1, row = 1)
    tsvButton.grid(column = 1, row = 2)
    mrpLabel.grid(column = 1, row = 3)
    mrpButton.grid(column = 1, row = 4)
    outputLabel.grid(column=1, row = 5)
    outputButton.grid(column = 1, row = 6)
    createButton.grid(column = 1, row = 7)
    exitButton.grid(column = 1, row = 8)

    window.mainloop()

if __name__ == "__main__":
    main()
