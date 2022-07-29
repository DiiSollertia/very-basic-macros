import requests
import os
import subprocess

def combineTS(url):
    counter = 0
    url = url[:146] + str(counter) + '.ts'
    #code = url[99:135]
    code = url[100:102]
    fileList = list()

    currentDir = os.path.dirname(__file__)
    path = os.path.join(currentDir, code)
    path = path.replace('/','\\')

    if not os.path.exists(path):
        os.makedirs(path)
        print('new folder created')
    else:
        return os.path.join(path,'all.mp4')

    GET = requests.get(url)
    while GET.status_code == 200:
        name = str(counter) + '.ts'
        file = open(os.path.join(path,name), 'wb')
        file.write(GET.content)
        file.close
        fileList.append(os.path.join(path,name))

        counter +=1
        url = url[:146] + str(counter) + '.ts'
        GET = requests.get(url)

    print(str(len(fileList)) + ' files created.')
    fileList = '+'.join(fileList)
    
    bashCommand =['copy','/b',fileList,os.path.join(path,'all.ts')]
    process = subprocess.Popen(bashCommand, stdout=subprocess.PIPE, text=True, shell=True)
    output, error = process.communicate()
    print('Command executed.')

    return os.path.join(path,'all.mp4')


vidURL = str(input('Video Request URL='))
vidOut = combineTS(vidURL)
audURL = str(input('Audio Request URL='))
audOut = combineTS(audURL)

