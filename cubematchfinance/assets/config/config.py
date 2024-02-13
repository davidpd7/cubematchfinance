import os
import json

def cfg_item(*items):
    data = Config.instance().data

    for key in items:
        data = data[key]

    return data

class Config:
    
    __current_dir = os.path.dirname(__file__) 
    __config_json_filename = "config.json"
    __instance = None

    @staticmethod
    def instance():
        if Config.__instance is None:
            Config()
        return Config.__instance

    def __init__(self):
        if Config.__instance is None:
            Config.__instance = self

            config_json_path = os.path.join(Config.__current_dir, Config.__config_json_filename)
            
            with open(config_json_path) as file:
                self.data = json.load(file)
        else:
            raise Exception("Config only can be instantiated once")