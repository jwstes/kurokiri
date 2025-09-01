from kuro import processInputToPrompt




if __name__ == "__main__":
    input = {
        "prompt" : "Write me out all the emails in this excel file",
        "files" : [
            "C:/Users/user/Desktop/kurokiri/testContent/test.docx",
            "C:/Users/user/Desktop/kurokiri/testContent/test.mp3", 
            "C:/Users/user/Desktop/kurokiri/testContent/test.pdf",
            "C:/Users/user/Desktop/kurokiri/testContent/test.pptx",
            "C:/Users/user/Desktop/kurokiri/testContent/test.xlsx",
            "C:/Users/user/Desktop/kurokiri/testContent/test.png"
        ]
    }

    finalPrompt = processInputToPrompt(input)

    file = open('prompt.txt', 'w', encoding="utf8")
    file.writelines(finalPrompt)
    file.close()