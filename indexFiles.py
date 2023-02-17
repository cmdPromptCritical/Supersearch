### crawls folder to produce index of all files and folders
import os
from datetime import datetime

now = datetime.now().strftime("%Y%m%d-%H:%M:%S")

### CONFIG
'''
    For the given path, get the List of all files in the directory tree 
'''     
rootFolders = [r'C:\auxDrive\work', 'gg'];
saveFile = r'fileIndex.txt'

### END CONFIG

def indexFiles(rootFolders: list, saveFile: str):
    myfile = open(saveFile, 'w', encoding='utf-8')
    for root in rootFolders:
        for path, subdirs, files in os.walk(root):
            for name in files:
                x = os.path.join(path, name)
                #print(x)
                myfile.write(f'{x}\n')
    
if __name__ == "__main__":
    indexFiles(rootFolders, saveFile)