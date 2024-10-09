from modules.opex import Opex
from modules.formatting import Formatting
import os


cwd = os.getcwd()
path=os.path.join(cwd, "data", "M_OPEX.xlsx")

opex = Opex(path)

opex.run()

formatting = Formatting(output="Opex.xlsx")

formatting.run()


