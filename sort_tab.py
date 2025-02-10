import win32clipboard
import mouse, keyboard
from time import sleep
from time import strftime
from random import random
import re
import json
import os

inventory = []
sorted_inventory = []
moves = 0

start_xy, end_xy, cell_width, cell_height, sort_index = None, None, None, None, None

def read_config():
    global start_xy, end_xy, cell_width, cell_height, sort_index
    basedir = os.path.dirname(__file__)
    with open(os.path.join(basedir, "config.json"), "r") as f:
        config = json.load(f)
        start_xy = config["start_xy"]
        end_xy = config["end_xy"]
        sort_index = config["sort_index"]
    print("loaded config.")
    print(json.dumps(config, indent=4))

def write_config():
    basedir = os.path.dirname(__file__)
    config = {
        "start_xy": start_xy,
        "end_xy": end_xy,
        "sort_index": sort_index
    }
    with open(os.path.join(basedir, "config.json"), "w") as f:
        json.dump(config, f, indent=4)
    print("config saved.")

def get_clipboard():
    try:
        win32clipboard.OpenClipboard()
    except Exception as e:
        sleep(0.05)
        win32clipboard.OpenClipboard()
    try:
        data = win32clipboard.GetClipboardData()
    except TypeError:
        data = None
    win32clipboard.CloseClipboard()
    return data

def clear_clipboard():
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.CloseClipboard()

def movetocell(x, y):
    real_x = start_xy[0] + (cell_width * x)
    real_y = start_xy[1] + (cell_height * y)
    #print(f"moving mouse to cell {x},{y} at {real_x}, {real_y} ")
    mouse.move(real_x, real_y)

def read_cell(x,y,i):
    sleep(0.025)
    clear_clipboard()
    sleep(0.025)
    keyboard.press_and_release('ctrl+c')
    sleep(0.025)
    contents = get_clipboard()
    sleep(0.025)
    cell = {
            "x" : x,
            "y" : y,
            "i" : i,
            "name": None,
            "raw": None,
            "item_class": None,
            "item_level": None,
            "waystone_tier": None,
            "waystone_drop_chance": None,
            "preferred_type": 0
        }
    if contents is not None:
        name = (contents.split("\n")[2]).strip()
        if name:
            cell['name'] = name
            print(f"Identified item: {name}")
        ilm = re.search(r"Item Level: (\d+)", contents)
        il = int(ilm.group(1)) if ilm else None
        if il:
            cell["item_level"] = il
        icm = re.search(r"Item Class: ([\w ]+)", contents)
        ic = icm.group(1) if icm else None
        if ic in sort_index:
            cell["preferred_type"] = sort_index[ic]
        cell["item_class"] = ic
        cell["raw"]        = contents
        if ic == "Waystones":
            tier_match = re.search(r"Waystone Tier: (\d+)", contents)
            wt = int(tier_match.group(1)) if tier_match else None
            drop_match = re.search(r"Waystone Drop Chance: \+(\d+)%", contents)
            dc = int(drop_match.group(1)) if drop_match else None
            cell["waystone_tier"] = wt
            cell["waystone_drop_chance"] = dc
    return cell

def swap_cell(cell, start=False):
    global moves
    moves+=1
    sleep(0.05)
    movetocell(cell["x"], cell["y"])
    sleep(0.05)
    if start:
        mouse.click()
        sleep(0.05)
    movetocell(cell["new_x"], cell["new_y"])
    sleep(0.05)
    mouse.click()
    sleep(0.05)

def recursive_swap(cell, start=False):
    if keyboard.is_pressed('alt'):
        print("alt pressed - exiting")
        return
    if "new_index" in cell and cell["item_class"] is not None:
        if cell["i"] == cell["new_index"]:
            #print(f"cell {cell['i']} is sorted.")
            return
        print(f"cell {cell['i']} ({cell['x']},{cell['y']}) -> {cell['new_index']} ({cell['new_x']},{cell['new_y']})")
        swap_cell(cell, start)
        cell["i"] = cell["new_index"]
        cell["x"] = cell["new_x"]
        cell["y"] = cell["new_y"]
        temp = inventory[cell["new_index"]]
        inventory[cell["new_index"]] = cell
        if temp:
            print(f"target cell contained item, recursing into cell {cell['new_index']}.")
            recursive_swap(temp)

def read_tab(size):
    global inventory
    global sorted_inventory
    global cell_width
    global cell_height
    cell_width  = (end_xy[0] - start_xy[0]) / (size-1)
    cell_height = (end_xy[1] - start_xy[1]) / (size-1)
    for i in range(size**2):
        if keyboard.is_pressed('alt'):
            print("alt pressed - exiting")
            break
        x = i // size
        y = i % size
        cell_position_x = start_xy[0] + (cell_width  * x)
        cell_position_y = start_xy[1] + (cell_height * y)
        mouse.move(cell_position_x, cell_position_y)
        cell = read_cell(x,y,i)
        inventory.append(cell)
    
    sorted_inventory = sorted( inventory, 
                                reverse=True,
                                key=lambda x: ( x["preferred_type"]       if x["preferred_type"]       is not None else 0,
                                                x["item_class"]           if x["item_class"]           is not None else "",
                                                x["waystone_tier"]        if x["waystone_tier"]        is not None else 0, 
                                                x["waystone_drop_chance"] if x["waystone_drop_chance"] is not None else 0,
                                                x["item_level"]           if x["item_level"]           is not None else 0,
                                                x["name"]                 if x["name"]                 is not None else ""
                                            ))
    for i in range(len(sorted_inventory)):
        sorted_inventory[i]["new_index"] = i
        sorted_inventory[i]["new_x"] = i // size
        sorted_inventory[i]["new_y"] = i % size
        inventory[sorted_inventory[i]["i"]]["new_index"] = i
        inventory[sorted_inventory[i]["i"]]["new_x"] = i // size
        inventory[sorted_inventory[i]["i"]]["new_y"] = i % size
        
    timestamp = strftime("%Y%m%d-%H%M%S")
    with open(f"inventory_{timestamp}.json", "w") as f:
        json.dump(sorted_inventory, f, indent=4)
    print("raw inventory saved.")

read_config()

print("listening for input...")
while(1):
    sleep(0.5)
    if keyboard.is_pressed('ctrl') and keyboard.is_pressed('5'):
        start_xy = mouse.get_position()
        print(f"start_xy set to {start_xy}")
        write_config()

    if keyboard.is_pressed('ctrl') and keyboard.is_pressed('6'):
        end_xy = mouse.get_position()
        print(f"end_xy set to {end_xy}")
        write_config()

    if keyboard.is_pressed('ctrl') and keyboard.is_pressed('1'):
        inventory=[]
        mouse.move(start_xy[0], start_xy[1])
        sleep(0.5)
        read_tab(12)
    if keyboard.is_pressed('ctrl') and keyboard.is_pressed('2'):
        inventory=[]
        mouse.move(start_xy[0], start_xy[1])
        sleep(0.5)
        read_tab(24)

    if keyboard.is_pressed('ctrl') and keyboard.is_pressed('3'):
        mouse.move(start_xy[0], start_xy[1])
        sleep(0.5)
        for i in range(len(inventory)):
            if keyboard.is_pressed('alt'):
                print("alt pressed - exiting")
                break
            recursive_swap(inventory[i], start=True)
        print(f"sorted in {moves} moves.")
    
