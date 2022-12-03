import os

files = os.listdir('./output')
files.remove('.gitkeep')
for file in files:
    os.remove('./output/' + file)