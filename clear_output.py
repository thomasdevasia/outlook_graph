import os

files = os.listdir('./output')
for file in files:
    os.remove('./output/' + file)