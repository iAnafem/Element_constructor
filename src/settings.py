import json
import os
from src.singleton import *


@singleton
class Settings:
    def __init__(self):
        with open(os.path.abspath('etc\\config.json'), encoding='utf-8') as data_file:
            settings = json.loads(data_file.read())
        self.groups = settings["groups"]
        self.scale_front = settings["scales"]["front"]
        self.scale_top = settings["scales"]["top"]
        self.scale_side = settings["scales"]["side"]
        self.holes_offset = settings["holes_offset"]
        self.autocad = os.path.abspath(settings["paths"]["Autocad"])
        self.excel = os.path.abspath(settings["paths"]["Excel"])
        self.workpath = (settings["paths"]["Workpath"])
