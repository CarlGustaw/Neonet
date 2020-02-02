import textract


def readDocx(path):
    text = textract.process(path)
    print(type(text))
