# Standard import
print('Importing packages ...')
import xlrd
import os
from pathlib import Path
print('- Package imported')

# Task 1: Write code to rename the files based on our dictionary

# Task 1.1: Pull video IDs and Titles from excel file
print('Pulling video IDs and Titles from Excel ...')

file_location = "/Users/admin/Desktop/Python_MacPro/organize_netflix/database.xlsx"
workbook = xlrd.open_workbook(file_location)
sheet_video = workbook.sheet_by_name('video-test')
video_titles = sheet_video.col_values(0, start_rowx=1)
video_ids = sheet_video.col_values(1, start_rowx=1)
video_ids_str = [str(int(id)) for id in video_ids]

# Task 1.2: Match IDs with Titles
print('Matching video IDs with Titles ...')
id_dict = {}
for i in range(len(video_ids_int)):
    id_dict[video_ids_int[i]] = video_titles[i]

# Task 1.3: Rename each file from its ID to its Title
print('Renaming files from IDs to Titles ...')

path = '/Users/admin/Desktop/.../organizeVideo/test'
os.chdir(path)
for f in os.listdir():
    for ids, titles in id_dict.items():
        video_name, suffix = os.path.splitext(f)
        if str(ids) == video_name:
            name = f'{titles}{suffix}'
            os.rename(f, name)

# Task 2: Write code to move the files into the folder of each genre

# Pull the genres data from Excel
print('Pulling genres from Excel ...')
sheet_genre = workbook.sheet_by_name('main-genres')
genres = sheet_genre.col_values(0, start_rowx=1)

# Match each file with its genre
print('Matching each file with its genre ...')

os.chdir(path)
genre_dict = {}
for i in genres:
    genre_dict[i] = []
for item in os.listdir(): 
    for genre, videos in genre_dict.items():
        if genre in item:
            videos.append(item)
def pick_directory(value):
    for genre, videos in genre_dict.items():
        for video in videos:
            if value == video:
                return genre
    return 'Others'

# Create the folders of genres and Move files accordingly
print('Creating folders of genres and Moving files ...')

os.chdir(path)
for item in os.scandir():
    # print(item)
    if item.is_dir():
        # print(item)
        continue
    filePath = Path(item)
    filePath_str = str(filePath)
    # print(filePath, type(filePath))
    directory = pick_directory(filePath_str)
    directoryPath = Path(directory)
    if directoryPath.is_dir() != True:
        directoryPath.mkdir()
    filePath.rename(directoryPath.joinpath(filePath))

print('Mission completed!!!')