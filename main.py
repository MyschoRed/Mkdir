import os
import py
with open("data.csv", "r") as file:
    data = file.readlines()
    folder_name = []
    for i in range(len(data)):
        folder_name.append(data[i])
for i in range(len(folder_name)):       
    os.mkdir(folder_name[i])