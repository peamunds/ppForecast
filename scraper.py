import numpy as np
import pandas as pd
import csv
from tkinter import *
from tkinter import filedialog
from PIL import Image, ImageTk

def getTsvFilepath():
    global tsvFileName
    tsvFileName = filedialog.askopenfilename(initialdir = "./",
                                          title = "Select a File")
    tsvLabel.configure(text=tsvFileName)
    if(tsvFileName and mrpFileName and outputFileName): # unlock create button if all paths aren't empty
        createButton.configure(state=ACTIVE)

def getMrpFilepath():
    global mrpFileName
    mrpFileName = filedialog.askopenfilename(initialdir = "./",
                                          title = "Select a File")
    mrpLabel.configure(text=mrpFileName)
    if(tsvFileName and mrpFileName and outputFileName): # unlock create button if all paths aren't empty
        createButton.configure(state=ACTIVE)

def getOutputFilepath():
    global outputFileName
    outputFileName = filedialog.asksaveasfilename(initialdir = "./",
                                          title = "Select a File",
                                          filetypes = (("Excel files",
                                                        "*.xlsx"),
                                                       ("all files",
                                                        "*.*")))
    outputLabel.configure(text=outputFileName)
    if(tsvFileName and mrpFileName and outputFileName): # unlock create button if all paths aren't empty
        createButton.configure(state=ACTIVE)


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
    window.destroy()


def main():
    global window
    window = Tk()
    window.title('Projection Creator')
    window.geometry("720x500")
    window.config(background = "white")

    bannerFrame = Frame(window, bg="red")
    inputFrame = Frame(window, bg="white")
    outputFrame = Frame(window, bg="white")
    tsvFrame = Frame(inputFrame,bg="white")
    mrpFrame = Frame(inputFrame, bg="white")
    outButFrame = Frame(outputFrame, bg="white")
    createFrame = Frame(outputFrame, bg="white")

    bannerFrame.pack(side="top",fill="both",expand=True)
    inputFrame.pack(fill="both",expand=True)
    outputFrame.pack(side="bottom",fill="both",expand=True)
    tsvFrame.pack(side="top", expand=True)
    mrpFrame.pack(side="bottom",expand=True)
    outButFrame.pack(side="top", expand=True)
    createFrame.pack(side="bottom", expand=True)

    img = ImageTk.PhotoImage(file = "pplogo.jpg")
    imgLabel = Label(bannerFrame, image=img).pack(expand=True, fill="both", padx=0,pady=0)

    labelBG = "#e0e0e0"
    labelFG = "black"
    buttonBG = "white"
    buttonFG = "black"

    global tsvLabel, mrpLabel,outputLabel, createButton
    tsvLabel = Label(tsvFrame,
                    text = "Please select a .tsv file with IMMI Projections",
                    width = 50, height = 1,
                    fg = labelFG, bg = labelBG, anchor="w")
    
    tsvButton = Button(tsvFrame,
                        text = "Browse",
                        command = getTsvFilepath,
                        fg=buttonFG,bg=buttonBG,bd=0, highlightthickness=0)

    mrpLabel = Label(mrpFrame,
                    text = "Please select an .xls file with MRP Projections",
                    width = 50, height = 1,
                    fg = labelFG, bg = labelBG, anchor="w")
    
    mrpButton = Button(mrpFrame,
                       text="Browse",
                       command=getMrpFilepath,
                       fg=buttonFG,bg=buttonBG,bd=0, highlightthickness=0)

    outputLabel = Label(outButFrame,
                    text = "Please select an .xlsx output file",
                    width = 50, height = 1,
                    fg = labelFG, bg = labelBG, anchor="w")

    outputButton = Button(outButFrame,
                          text="Browse",
                          command=getOutputFilepath,
                          fg=buttonFG,bg=buttonBG,bd=0, highlightthickness=0)
    
    createButton = Button(createFrame,
                        text="Create Forecast",
                        command=createForecast,
                        fg=buttonFG,bg=buttonBG,state=DISABLED,bd=0,highlightthickness=0)


    tsvLabel.pack(side="left",expand=True)
    tsvButton.pack(side="right",expand=True)
    mrpLabel.pack(side="left",expand=True)
    mrpButton.pack(side="right",expand=True)
    outputLabel.pack(side="left",expand=True)
    outputButton.pack(side="right",expand=True)
    createButton.pack(expand=True)
    # exitButton.grid(  column = 1, row = 8)

    window.mainloop()

if __name__ == "__main__":
    main()
