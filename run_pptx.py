from pptx import Presentation
import codecs
prs = Presentation()
text_path = "D:/python/수련회찬양.txt"
f = codecs.open(text_path, "r", "utf-8")

lines = f.readlines()
for line in lines:
    print(line.strip())
    break