import os
import numpy as ny
import cv2

from PIL import Image

names=[]
paths=[]

for users in os.listdir("C:\\Users\\Uthayakumar Thenujan\\OneDrive\\Desktop\\Python\\Face\\dataset"):
    names.append(users)

for name in names:
    for image in os.listdir("C:\\Users\\Uthayakumar Thenujan\\OneDrive\\Desktop\\Python\\Face\\dataset\\{}".format(name)):
        path_string= os.path.join("C:\\Users\\Uthayakumar Thenujan\\OneDrive\\Desktop\\Python\\Face\\dataset\\{}".format(name),image)
        paths.append(path_string)

faces =[]
ids =[]

for img_path in paths:
    image =Image.open(img_path).convert("L")

    imgNp = ny.array(image, "uint8")

    faces.append(imgNp)

    id =int(img_path.split("\\")[9].split("_")[0])

    ids.append(id)

ids=ny.array(ids)

trainer = cv2.face.LBPHFaceRecognizer_create()

trainer.train(faces,ids)

trainer.write("C:\\Users\\Uthayakumar Thenujan\\OneDrive\\Desktop\\Python\\Face\\training.yml")
print("Sucess!!")