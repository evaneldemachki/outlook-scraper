import core
import json
import pandas

def init():
    config = _sess.config
    #sess.update()

    index = {}
    for account in config["index"]:
        entry = {
            "object": None,
            "inbox": {
                "live": {},
                "archive": {}
            }
        }

        live_data = _sess.load(account)
        arch_data = _sess.load_archive(account)
        # generate archive index for live data model
        arch_index = []
        for id in live_data['entry_id']:
            match = arch_data.loc[arch_data['entry_id'] == id]['category']
            if match.shape[0] > 0:
                arch_index.append(match.iloc[0])
            else:
                arch_index.append(False)
        # add archive_index column to live data
        arch_index = pandas.Series(arch_index)
        live_data['archive_index'] = arch_index

        # rename columns to numbered
        #live_data.columns = [0,1,2,3]
        entry["inbox"]["live"] = {
            "object": None,
            "data": live_data
        }
        entry["inbox"]["archive"] = {
            "object": None,
            "data": arch_data
        }

        index[account] = entry

    # generate category cache entry
    categories = {}
    for category in config["categories"]:
        categories[category] = None
    # look for existing color configuration
    for category, color in config["gui"]["category_colors"].items():
        if category in categories: # assure color entry matches category entry
            # add color to entry category
            categories[category] = color
    # fill in non-existent entries with randomly generated colors
    for category in config["categories"]:
        if category not in config["gui"]["category_colors"]:
            rgb_color = tuple(random.randint(0,255) for i in range(0,3))
            hex_color = '#%02x%02x%02x' % (*rgb_color,)
            categories[category] = hex_color

    cache = {
        "index": index,
        "categories": categories
    
    }
    return cache

def editconfig(section, key, value):
    with open("config.json", "r") as f:
        conf_file = json.load(f)

    if section == "color":
        # convert QColor to HEX
        hex_color = value.name()
        conf_file["gui"]["category_colors"][key] = hex_color
    else:
        msg = "invalid configuration modifier was called"
        raise UiRuntimeError(msg)

def sess():
    return _sess

_sess = core.Session()