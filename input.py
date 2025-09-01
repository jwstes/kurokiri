from toolbox import docxExtract, pdfExtract, pptxExtract, xlsxExtract
import os
import json

def processInputToPrompt(input):
    inputFiles = input['files']
    finalPrompt = f"User Prompt: {input['prompt']}\n\n"
    if len(inputFiles) > 0:
        finalPrompt += "User Uploaded Files: \n\n"
        for file in inputFiles:
            _, ext = os.path.splitext(file)
            ext = ext.lower()
            md = ""
            if ext == '.pdf':
                md += f"PDF File: {file}\n"
                md += pdfExtract.main(file)
            elif ext == '.docx':
                md += f"DOCX File: {file}\n"
                md += docxExtract.main(file)
            elif ext == '.xlsx':
                md += f"XLSX File: {file}\n"
                md += xlsxExtract.main(file)
            elif ext == '.pptx':
                md += f"PPTX File: {file}\n"
                md += pptxExtract.main(file)
            else:
                print(f"Unsupported file type ({ext}): {file}")
            finalPrompt += f"{md}\n\n\n\n"
    return finalPrompt


input = {
    "prompt" : "Write me out all the emails in this excel file",
    "files" : [
        # "C:/Users/otter/Desktop/Kurogi/test.pdf",
        # "C:/Users/otter/Desktop/Kurogi/test.docx",
        "C:/Users/otter/Desktop/Kurogi/test.xlsx",
        # "C:/Users/otter/Desktop/Kurogi/test.pptx",
    ]
}

finalPrompt = processInputToPrompt(input)




testPrompt = {
    "stream": True, 
    "messages": 
    [{"role": "user", "content": finalPrompt}], 
    "max_tokens": 4096, "temperature": 0.5, "chat_template_kwargs": 
    {"enable_thinking": True}
    }

file = open('prompt3.json', 'w', encoding="utf8")
file.writelines(json.dumps(testPrompt))
file.close()