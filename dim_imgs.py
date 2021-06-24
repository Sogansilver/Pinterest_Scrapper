from cv2 import cv2
import pathlib
import os

ruta_completa = os.getcwd()
dim = 256
# dim = 64
entries = os.listdir(ruta_completa+"/img_native/")
for entry in entries:
    paths = ruta_completa+"/img_native/"+entry
    if os.path.exists(ruta_completa+"/img_final/"+entry) == True:
       print("El archivo ya existe en img_final") 
    else:
        print("El archivo no existe en img_final") 
        try:
            imgs=cv2.imread(paths)
            if imgs.shape[0] > imgs.shape[1]:
                x =round(imgs.shape[1]/imgs.shape[1]*dim)
                y =round(imgs.shape[0]/imgs.shape[1]*dim)
            else:
                x =round(imgs.shape[1]/imgs.shape[0]*dim)
                y =round(imgs.shape[0]/imgs.shape[0]*dim)
            redim = (x,y)
            sized_img = cv2.resize(imgs,redim)
            cropped_img = sized_img[0:dim,0:dim]
            cv2.imwrite(ruta_completa +"/img_final/"+entry,cropped_img) 
        except:
            print("Error")
