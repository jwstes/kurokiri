from kuro import processInputToPrompt
from largemodels.qwen3_streamer import init_llm, stream_prompt

if __name__ == "__main__":
    input = {
        "prompt" : "List out all the emails in this excel file.",
        "files" : [
            # "C:/Users/user/Desktop/kurokiri/testContent/test.docx",
            # "C:/Users/user/Desktop/kurokiri/testContent/test.mp3", 
            # "C:/Users/user/Desktop/kurokiri/testContent/test.pdf",
            # "C:/Users/user/Desktop/kurokiri/testContent/test.pptx",
            "C:/Users/user/Desktop/kurokiri/testContent/test.xlsx",
            # "C:/Users/user/Desktop/kurokiri/testContent/x.jpg"
        ]
    }

    finalPrompt = processInputToPrompt(input)

    init_llm()
    for chunk in stream_prompt(finalPrompt, hide_thinking=False, max_new_tokens=4096):
        print(chunk, end="", flush=True)
    print()

    # file = open('prompt.txt', 'w', encoding="utf8")
    # file.writelines(finalPrompt)
    # file.close()