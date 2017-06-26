# -*- coding: utf-8 -*-
"""
locations.py

Variables containing the .pak file and internal location of specific data sources.
"""

gamedata_pak = 'GameData.pak'
mech_file = (gamedata_pak, 'Libs/Items/Mechs/Mechs.xml')
weapon_file = (gamedata_pak, 'Libs/Items/Weapons/Weapons.xml')
equipment_file = (gamedata_pak, 'Libs/Items/Modules/Equipment.xml')
consumable_file = (gamedata_pak, 'Libs/Items/Modules/Consumables.xml')
engine_file = (gamedata_pak, 'Libs/Items/Modules/Engines.xml')
masc_file = (gamedata_pak, 'Libs/Items/Modules/MASC.xml')
skill_tree_file = (gamedata_pak, 'Libs/MechPilotTalents/MechSkillTreeNodes.xml')
skill_tree_layout_file = (gamedata_pak, 'Libs/MechPilotTalents/MechSkillTreeNodesDisplay.xml')

english_pak = 'Localized/English.pak'
localization_file = (english_pak, 'Languages/TheRealLoc.xml')