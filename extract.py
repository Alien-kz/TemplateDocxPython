import argparse
import os
import docx2txt
from docx import Document

DEBUG = False

START = 1
END = 2

def clean(text):
    return ' '.join(text.replace("_", "").strip().split())

def debug(state, source_text, found_text):
    if DEBUG and state > 0:
        info = str(state)
        if source_text:
            info +=  "|" + source_text.replace("\n", " ")
        if found_text:
            info += " >>> " + found_text.replace("\n", " ")
        print(info)

def step(state, source_text, start, end):
    found_text = source_text
    new_state = 00
    if start in found_text:
        new_state = new_state | START
        words = found_text.split(start)
        if len(words) > 0 and words[-1]:
            found_text = clean(words[-1])
        else:
            found_text = ""

    if end in found_text:
        new_state = new_state | END
        words = found_text.split(end)
        if len(words) > 0 and words[0]:
            found_text = clean(words[0])
        else:
            found_text = ""

    debug(new_state, source_text, found_text)
    return new_state, found_text

def run(file_path, start, end, shift):
    source_texts = docx2txt.process(file_path)
    source_texts = source_texts.split("\n")

    texts = []
    state = 0 # start/end: 0 - 0/0 , 1 - 1/0, 2 - 0/1, 3 - 1/1
    started = False
    for source_text in source_texts:
        state, found_text = step(state, source_text, start, end)
        if found_text:
            if START & state:
                texts = [found_text]
            elif started:
                texts.append(found_text)
        if START & state:
            started = True
        if END & state and texts:
            break

    if shift:
        texts = [texts[shift]]
    return '\n'.join(texts).strip()

def main(source, start, end, shift, destination):
    files = [file for file in sorted(os.listdir(source)) if file.endswith(".docx")]
    print(source)
    print(len(files), ":", files[:5])
    print()

    print("START: \"" + start + "\"")
    print("END: \"" + end + "\"")
    print()

    result = []
    for file in files:
        file_path = os.path.join(source, file)
        found_text = run(file_path, start, end, shift)

        file = os.path.splitext(file)[0]
        print("FILE:", file)
        print("TEXT:", found_text)
        result.append(file + ",\"" + found_text + "\"")
        print()
        if DEBUG:
            break
    header = "file,text\n"
    result = "\n".join(result) + "\n"
    print(destination)
    with open(destination, "w") as f:
        f.write(header)
        f.write(result)

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("source", type=str, help="[input] directory")
    parser.add_argument("start", type=str, help="[input] start text")
    parser.add_argument("end", type=str, help="[input] end text")
    parser.add_argument("destination", type=str, help="[output] csv")
    parser.add_argument("--shift", type=int, help="[input] end text", default=0)
    parser.add_argument("--debug", type=str, help="[output] csv")
    args = parser.parse_args()
    if args.debug:
        DEBUG = True
    main(args.source, args.start, args.end, args.shift, args.destination)