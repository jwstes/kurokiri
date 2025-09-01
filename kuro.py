import os
import sys
import io
import json
import multiprocessing as mp
from multiprocessing import Process, Queue
from queue import Empty


def pdfWorker(result_queue, inputPath):
    try:
        from toolbox import pdfExtract
        md = pdfExtract.main(inputPath)
        result_queue.put(md)
    except Exception as e:
        result_queue.put(("err", f"{type(e).__name__}: {e}"))

def docxWorker(result_queue, inputPath):
    try:
        from toolbox import docxExtract
        md = docxExtract.main(inputPath)
        result_queue.put(md)
    except Exception as e:
        result_queue.put(("err", f"{type(e).__name__}: {e}"))

def xlsxWorker(result_queue, inputPath):
    try:
        from toolbox import xlsxExtract
        md = xlsxExtract.main(inputPath)
        result_queue.put(md)
    except Exception as e:
        result_queue.put(("err", f"{type(e).__name__}: {e}"))

def pptxWorker(result_queue, inputPath):
    try:
        from toolbox import pptxExtract
        md = pptxExtract.main(inputPath)
        result_queue.put(md)
    except Exception as e:
        result_queue.put(("err", f"{type(e).__name__}: {e}"))

def audioWorker(result_queue, inputPath):
    try:
        from toolbox import audioExtract
        md = audioExtract.main(inputPath)
        result_queue.put(md)
    except Exception as e:
        result_queue.put(("err", f"{type(e).__name__}: {e}"))

def imageWorker(result_queue, inputPath):
    try:
        from toolbox import imageExtract
        md = imageExtract.main(inputPath)
        result_queue.put(md)
    except Exception as e:
        result_queue.put(("err", f"{type(e).__name__}: {e}"))


def multiProcessStarter(worker, inputPath, timeout=300):
    rQ = mp.Queue()
    p = mp.Process(target=worker, args=(rQ, inputPath))
    p.start()
    try:
        result = rQ.get(timeout=timeout)
    except Empty:
        alive = p.is_alive()
        p.terminate()
        p.join()
        raise RuntimeError(
            f"{worker.__name__} {'hung' if alive else 'exited'} without producing a result for {inputPath}"
        )
    finally:
        rQ.close()
        rQ.join_thread()

    p.join()
    if p.exitcode != 0:
        raise RuntimeError(
            f"{worker.__name__} exited with code {p.exitcode} for {inputPath}"
        )
    return result

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
                md += multiProcessStarter(pdfWorker, file)
            elif ext == '.docx':
                md += f"DOCX File: {file}\n"
                md += multiProcessStarter(docxWorker, file)
            elif ext == '.xlsx':
                md += f"XLSX File: {file}\n"
                md += multiProcessStarter(xlsxWorker, file)
            elif ext == '.pptx':
                md += f"PPTX File: {file}\n"
                md += multiProcessStarter(pptxWorker, file)
            elif ext == '.mp3':
                md += f"Audio File: {file}\n"
                md += multiProcessStarter(audioWorker, file)
            elif ext == '.png' or ext == '.jpg' or ext == '.jpeg':
                md += f"Image {file} Details:\n"
                md += multiProcessStarter(imageWorker, file)
            else:
                print(f"Unsupported file type ({ext}): {file}")
            finalPrompt += f"{md}\n\n\n\n"
    return finalPrompt




