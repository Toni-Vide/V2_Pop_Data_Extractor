import regex as re
import pandas as pd
from pathlib import Path  # used to read the data in the file
import tkinter as tk
from tkinter import filedialog  # used to open a dialogue in order to get the save game
import os  # used to get filename

# opens a dialogue that prompts the user to select a file
root = tk.Tk()
root.withdraw()
path = filedialog.askopenfilename()

# checks if the file is indeed a v2 save
if path[-2:] != 'v2':
    raise Exception('Warning the selected file is not a v2 save-game as identified with the v2 suffix. This operation '
                    'will Terminate.')


filename = os.path.basename(path).split('/')[-1]  # gets the file name
data = Path(path).read_text()


accumulator = []  # holds the data that we will at the end commit to the df
pop_sequence = 1  # sets the first number in pop sequence in order to give an ID to the pop
province = re.findall(r'(\d+)=(\s*{(?:[^{}]+|(?2))*})', data)  # extracts the data of the individual provinces

# ^(?![a-z])(\d+) = { }

for data in province:

    province_id = data[0]

    #  for whatever reason some tags are called d01 etc... which is messing with the regex
    if province_id == '01':
        break

    province_data = data[1]
    province_name = re.findall(r'name="([^"]+)"', province_data)[0]

    #  In case the land is unclaimed
    try:
        province_owner = re.findall(r'owner="([^"]+)"', province_data)[0]
    except IndexError:
        province_owner = "Unclaimed Land"

    #  In case the land is unclaimed
    try:
        province_controller = re.findall(r'controller="([^"]+)"', province_data)[0]
    except IndexError:
        province_owner = "Unclaimed Land"

    province_pops = re.findall(r'(\w+)=\s+{\s+id=(.*)\s+size=(.*)\s+(\w+)=(\w+)\s+money=(.*)\s+', province_data)

    for pop in range(0, len(province_pops)):
        accumulator.append([str(pop_sequence),  # unique ID
                            province_id,  # province id
                            province_name,  # province name
                            province_owner,  # province owner
                            province_controller,  # province controller
                            str(province_pops[pop][1]),  # pop ID
                            province_pops[pop][0],  # profession
                            province_pops[pop][2],  # size
                            str(province_pops[pop][3]),  # culture
                            province_pops[pop][4],  # religion
                            round(float(province_pops[pop][5]), 2)  # money
                            ])
        pop_sequence += 1


population = pd.DataFrame(accumulator,
                          columns=[
                            'pop_sequence',
                            'province_id',
                            'province_name',
                            'province_owner',
                            'province_controller',
                            'pop_id',
                            'pop_profession',
                            'pop_size',
                            'pop_culture',
                            'pop_religion',
                            'pop_money'
                            ])

# asks the user where to save the file
savepath = filedialog.askdirectory()
savepath = savepath + '\\'

population.to_excel(savepath+filename+'.xlsx', sheet_name='Data')
