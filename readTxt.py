
f = open("D:/python/PPTX/test.txt")
lines = f.readlines()
for line in lines:
    if(line.startswith("##")):
        print("title:"+line[2:].strip())
    elif(line.startswith("#")):
        print("next:")
    else:
        print(line.strip())


f.close()