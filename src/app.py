from modules.opex import Opex
from modules.formatting import Formatting
import os

path=r"C:\Users\anderson.bones\Desktop\opex-fs\src\data\M_OPEX.xlsx"
opex = Opex(path)

opex.run()

formatting = Formatting(output="Opex.xlsx")

formatting.run()

cwd = os.getcwd()
