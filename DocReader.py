import textract


def readDoc(path):
    text = textract.process(path)
    print(type(text))
