import mouse, keyboard
from time import sleep
from time import strftime
import win32clipboard # type: ignore
from random import random
import re
import json

inventory = []
moves = 0

sort_index = {
    "Waystones": 4,
    "Tablet": 2,
    "Stackable Currency": 1,
}

start_xy = (35, 145)
end_xy = (625, 725)
cell_width  = (end_xy[0] - start_xy[0]) / 11
cell_height = (end_xy[1] - start_xy[1]) / 11

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
    mouse.move(start_xy[0] + (cell_width * x), start_xy[1] + (cell_height * y))

def read_cell(x,y,i):
    sleep(0.025)
    clear_clipboard()
    sleep(0.025)
    keyboard.press_and_release('ctrl+c')
    sleep(0.025)
    contents = get_clipboard()
    sleep(0.025)
    if contents is not None:
        match = re.search(r"Item Class: ([\w ]+)", contents)
        ic = match.group(1) if match else None
        if ic in sort_index:
            ti = sort_index[ic]
        else:
            ti = 0
        cell = {
            "x" : x,
            "y" : y,
            "i" : i,
            "raw": contents,
            "waystone_tier": None,
            "waystone_drop_chance": None,
            "item_class": ic,
            "type_index": ti
        }
        if ic == "Waystones":
            tier_match = re.search(r"Waystone Tier: (\d+)", contents)
            wt = int(tier_match.group(1)) if tier_match else None
            drop_match = re.search(r"Waystone Drop Chance: \+(\d+)%", contents)
            dc = int(drop_match.group(1)) if drop_match else None
            cell["waystone_tier"] = wt
            cell["waystone_drop_chance"] = dc
        return cell
    else:
        return {
            "x" : x,
            "y" : y,
            "i" : i,
            "raw": None,
            "item_class": None,
            "waystone_tier": None,
            "waystone_drop_chance": None,
            "type_index": 0
        }

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
            recursive_swap(temp)

print("listening for input...")
while(1):
    sleep(0.5)
    if keyboard.is_pressed('ctrl') and keyboard.is_pressed('1'):
        mouse.move(start_xy[0], start_xy[1])
        sleep(0.5)

        inventory = []
        waystones = []

        for i in range(144):
            if keyboard.is_pressed('alt'):
                print("alt pressed - exiting")
                break
            x = i // 12
            y = i % 12
            cell_position_x = start_xy[0] + (cell_width * x)
            cell_position_y = start_xy[1] + (cell_height * y)
            mouse.move(cell_position_x, cell_position_y)
            cell = read_cell(x,y,i)
            inventory.append(cell)
        
        waystones = sorted(inventory, key=lambda x: ( x["type_index"] if x["type_index"] is not None else 0,
                                                      x["item_class"] if x["item_class"] is not None else "",
                                                      x["waystone_tier"] if x["waystone_tier"] is not None else 0, 
                                                      x["waystone_drop_chance"] if x["waystone_drop_chance"] is not None else 0), 
                                                      reverse=True)
        for i in range(len(waystones)):
            waystones[i]["new_index"] = i
            waystones[i]["new_x"] = i // 12
            waystones[i]["new_y"] = i % 12
            inventory[waystones[i]["i"]]["new_index"] = i
            inventory[waystones[i]["i"]]["new_x"] = i // 12
            inventory[waystones[i]["i"]]["new_y"] = i % 12

        timestamp = strftime("%Y%m%d-%H%M%S")
        with open(f"inventory_{timestamp}.json", "w") as f:
            json.dump(inventory, f, indent=4)
        print("raw inventory saved.")
        #with open(f"waystones_{timestamp}.json", "w") as f:
        #    json.dump(waystones, f, indent=4)
        #print("sorted inventory saved.")

    if keyboard.is_pressed('ctrl') and keyboard.is_pressed('2'):
        already_sorted_indexes = []
        mouse.move(start_xy[0], start_xy[1])
        sleep(0.5)
        for i in range(len(inventory)):
            recursive_swap(inventory[i], start=True)
        print(f"sorted in {moves} moves.")
        