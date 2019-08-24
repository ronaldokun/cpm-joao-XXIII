import os, sys

ROOT = os.path.abspath(os.path.join(os.path.dirname("__file__"), ".."))
sys.path.insert(0, ROOT)

AUTH = f"{ROOT}/credentials/drive.json"


import cpm
