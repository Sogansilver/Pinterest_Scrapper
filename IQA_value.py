import numpy as numpy
import skimage
from skimage import io, img_as_float
import imquality.brisque as brisque
import os

ruta_completa = os.getcwd()

entries = os.listdir(ruta_completa+"/img_final/")
for entry in entries:
    img_native = img_as_float(io.imread(ruta_completa +"/img_native/"+entry,as_gray=True))
    img_final = img_as_float(io.imread(ruta_completa +"/img_final/"+entry,as_gray=True))

    score_native = brisque.score(img_native)
    score_final = brisque.score(img_final)
    print("El archivo: "+entry+" posee un score nativo de: "+str(score_native)+", y un score final de: "+str(score_final))