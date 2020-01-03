from modules import interlingua_endocytosis
from modules import gui


path = gui()
if path['content']:
    interlingua_endocytosis(path['const'], path['add'], path['save'])
