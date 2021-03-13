import os
import re
fileType = '.xlsx'
path = '.'
newfolder =os.path.join(path,'00renamedfiles')
os.mkdir(newfolder)

for dir,folder,filenames in os.walk(path):
    for filename in filenames:
        if filename.endswith(fileType):
            studentid = re.findall(r'\((.*?)\)', dir)
            if studentid:
                studentid = studentid[0]
                newname = studentid+fileType
                os.rename(os.path.join(dir,filename),os.path.join(newfolder,newname))
