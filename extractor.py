# -*- coding: utf-8 -*-
"""
extractor.py

Extracts MWO game data to a more accessible format.
"""

import os
import zipfile
import winreg
from lxml import etree as ET
import json

# Ensure wdir is current
abs_path = os.path.abspath(__file__)
dir_name = os.path.dirname(abs_path)
os.chdir(dir_name)

from locations import gamedata_pak, weapon_file, mech_file, equipment_file, english_pak, localization_file

MWO_REG_KEY = 'SOFTWARE\\WOW6432Node\\Piranha Games\\Mechwarrior Online\\Production\\Install'

def get_game_dir():
    """
    Return MWO game data directory based on Windows registry key.
    """
    key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, MWO_REG_KEY)
    install_dir = winreg.QueryValueEx(key, "INSTALL_LOCATION")[0]
    game_dir = os.path.join(install_dir, 'Game')
    return game_dir

def map_localizations():
    """
    Return dictionary of localized text by tag string.
    """
    localizations = {}

    localedata_path = os.path.join(get_game_dir(), english_pak)
    with zipfile.ZipFile(localedata_path, 'r') as localedata:
        with localedata.open(localization_file[1]) as localization_data:
            LOC_TREE = ET.iterparse(localization_data)

            # Strip Excel namespaces from the XML
            for _, el in LOC_TREE:
                if '}' in el.tag:
                    el.tag = el.tag.split('}', 1)[1]

            LOC_ROOT = LOC_TREE.root

            ns = {
                'o':'urn:schemas-microsoft-com:office:office',
                'x':'urn:schemas-microsoft-com:office:excel',
                'ss':'urn:schemas-microsoft-com:office:spreadsheet',
                'ns':'urn:schemas-microsoft-com:office:spreadsheet'
            }

            locale_nodes = LOC_ROOT.xpath('//Row/Cell[@ss:Index]/Data', namespaces=ns)
            for locale_node in locale_nodes:
                locale_tag = locale_node.text.lower()
                locale_row = locale_node.getparent().getparent()
                locale_text = locale_row.getchildren()[1][0].text
                localizations[locale_tag] = locale_text

    return localizations


def map_weapons(localizations):
    """
    Return dictionary of weapons by ID.
    """
    weapons = {}

    gamedata_path = os.path.join(get_game_dir(), gamedata_pak)
    with zipfile.ZipFile(gamedata_path, 'r') as gamedata:
        with gamedata.open(weapon_file[1]) as weapon_data:
            WEAPON_TREE = ET.parse(weapon_data)
            WEAPON_ROOT = WEAPON_TREE.getroot()

            # TODO: need to use locale transform
            for weapon in WEAPON_ROOT.iter('Weapon'):

                # Skip Artemis weapons with identical base stats
                if 'InheritFrom' in weapon.attrib:
                    continue

                weapon_stats = weapon.find('WeaponStats')
                weapon_loc = weapon.find('Loc')

                weapon_id = weapon.attrib['id']
                weapon_name = localizations[weapon_loc.attrib['nameTag'][1:].lower()]
                weapon_type = weapon_stats.attrib['type']
                weapon_faction = weapon.attrib['faction']

                weapons[weapon_id] = {
                    # Weapon details
                    'id': weapon_id,
                    'name':  weapon_name,
                    'type': weapon_type,
                    'faction': weapon_faction,

                    # Equipment stats
                    'health': weapon_stats.attrib['Health'],
                    'weight': weapon_stats.attrib['tons'],
                    'slots': weapon_stats.attrib['slots'],

                    # Weapon stats
                    'damage': weapon_stats.attrib['damage'],
                    'heat': weapon_stats.attrib['heat'],
                    'duration': weapon_stats.attrib['duration'],
                    'cooldown': weapon_stats.attrib['cooldown'],

                    # Weapon behavior
                    'velocity': weapon_stats.attrib['speed'],
                    'count': weapon_stats.attrib['numFiring'],
                    'volleySize': weapon_stats.attrib['volleysize'] if 'volleysize' in weapon_stats.attrib else '1',
                    'volleyDelay': weapon_stats.attrib['volleydelay'] if 'volleydelay' in weapon_stats.attrib else '0',

                    'minRange': weapon_stats.attrib['minRange'],
                    'optRange': weapon_stats.attrib['longRange'],
                    'maxRange': weapon_stats.attrib['maxRange'],

                    # Ghost heat (group 0 does not exist)
                    'penalty': weapon_stats.attrib['heatpenalty'] if 'heatpenalty' in weapon_stats.attrib else '0',
                    'penaltyLimit': weapon_stats.attrib['minheatpenaltylevel'] if 'minheatpenaltylevel' in weapon_stats.attrib else '0',
                    'penaltyGroup': weapon_stats.attrib['heatPenaltyID'] if 'heatPenaltyID' in weapon_stats.attrib else '0'
                }

    return weapons


def map_mechs(localizations):
    """
    Return dictionary of mechs by ID.
    """
    mechs = {}

    gamedata_path = os.path.join(get_game_dir(), gamedata_pak)
    with zipfile.ZipFile(gamedata_path, 'r') as gamedata:
        with gamedata.open(mech_file[1]) as mech_data:
            MECH_TREE = ET.parse(mech_data)
            MECH_ROOT = MECH_TREE.getroot()

            # TODO: need to use locale transform
            for mech in MECH_ROOT.iter('Mech'):
                # Skip lame escort mech
                mech_name = mech.attrib['name']
                if 'escort' in mech_name:
                    continue

                mech_id = mech.attrib['id']
                mech_chassis = mech.attrib['chassis']

                # Apply localizations
                mech_full_name = localizations[mech_name]
                mech_chassis = localizations[mech_chassis]
                mech_variant = localizations[mech_name + '_short']

                mechs[mech_id] = {
                    'id': mech_id,
                    'name': mech_full_name,
                    'chassis': mech_chassis,
                    'variant': mech_variant
                }

    return mechs


def map_equipment(localizations):
    """
    Return dictionary of equipment by ID.
    """
    equipment = {}

    gamedata_path = os.path.join(get_game_dir(), gamedata_pak)
    with zipfile.ZipFile(gamedata_path, 'r') as gamedata:
        with gamedata.open(equipment_file[1]) as equipment_data:
            EQUIPMENT_TREE = ET.parse(equipment_data)
            EQUIPMENT_ROOT = EQUIPMENT_TREE.getroot()

            # Note: child elements named "Module", not "Equipment"
            for module in EQUIPMENT_ROOT.iter('Module'):
                module_id = module.attrib['id']
                module_loc = module.find('Loc')
                module_name = localizations[module_loc.attrib['nameTag'][1:].lower()]

                equipment[module_id] = {
                    'id': module_id,
                    'name': module_name
                }

    return equipment


def map_heat_sinks(localizations):
    """
    Return dictionary of heat_sinks by ID.
    """
    heat_sinks = {}

    gamedata_path = os.path.join(get_game_dir(), gamedata_pak)
    with zipfile.ZipFile(gamedata_path, 'r') as gamedata:
        with gamedata.open(equipment_file[1]) as equipment_data:
            EQUIPMENT_TREE = ET.parse(equipment_data)
            EQUIPMENT_ROOT = EQUIPMENT_TREE.getroot()
            HEAT_SINK_NODES = EQUIPMENT_ROOT.xpath('//Module[@CType="CHeatSinkStats"]')

            # Note: child elements named "Module", not "Equipment"
            for hs in HEAT_SINK_NODES:
                hs_id = hs.attrib['id']
                hs_loc = hs.find('Loc')
                hs_name = localizations[hs_loc.attrib['nameTag'][1:].lower()]
                hs_faction = hs.attrib['faction']

                hs_stats = hs.find('HeatSinkStats')
                hs_capacity = hs_stats.attrib['heatbase']
                hs_internal_delta = hs_stats.attrib['engineCooling']
                hs_external_delta = hs_stats.attrib['cooling']

                heat_sinks[hs_id] = {
                    'id': hs_id,
                    'name':  hs_name,
                    'faction': hs_faction,
                    'capacity': hs_capacity,
                    'internalDelta': hs_internal_delta,
                    'externalDelta': hs_external_delta
                }

    return heat_sinks


def export_json(filename, obj):
    if not os.path.exists(os.path.dirname(filename)):
        try:
            os.makedirs(os.path.dirname(filename))
        except OSError as exc: # Guard against race condition
            if exc.errno != errno.EEXIST:
                raise

    with open(filename, 'w') as json_file:
        json_file.write(json.dumps(obj, indent=2))


if __name__ == '__main__':
    # Extract data
    localizations = map_localizations()
    weapons = map_weapons(localizations)
    mechs = map_mechs(localizations)
    equipment = map_equipment(localizations)
    heat_sinks = map_heat_sinks(localizations)

    locale_json_file = 'out/locale.json'
    weapon_json_file = 'out/weapons.json'
    mech_json_file = 'out/mechs.json'
    equipment_json_file = 'out/equipment.json'
    heat_sink_json_file = 'out/heatsinks.json'

    export_json(locale_json_file, localizations)
    export_json(weapon_json_file, weapons)
    export_json(mech_json_file, mechs)
    export_json(equipment_json_file, equipment)
    export_json(heat_sink_json_file, heat_sinks)
