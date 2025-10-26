import os


def check(path):
    file_name, extention = os.path.splitext(path)
    output = file_name + ".txt"
    print(path)
    print(output)
    print(file_name)
    print(extention)


check("somefile.mp3")
