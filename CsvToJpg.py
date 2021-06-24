import pandas as pd
import urllib.request
import os
import pathlib

ruta_completa = os.getcwd()

def url_to_jpg(i,url,file_path):
  try:
    files = url
    files1= files[-5:-37:-1]
    filename = str(i+2)+"_"+files1[::-1]+".jpg"
  
    file = pathlib.Path(ruta_completa +"/img_native/"+filename)
    if file.exists ():
        print (str(i+2)+". El archivo "+filename+" ya existe en nativo")
    else:
        full_path = '{}{}'.format(file_path,filename)
        urllib.request.urlretrieve(url,full_path)
        print('{} se ha guardado en nativo.'.format(filename))
    return None
  except:
    print("Error")
  
FILENAME = ruta_completa + "/imgsUrls.csv"
FILE_PATH = ruta_completa + "/img_native/"

urls = pd.read_csv(FILENAME)

for i,url in enumerate(urls.values):
  url_to_jpg(i,url[0],FILE_PATH)