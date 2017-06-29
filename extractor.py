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

from locations import gamedata_pak, weapon_file, mech_file, english_pak, localization_file

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


def map_mechs():
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
                mech_id = mech.attrib['id']
                mech_name = mech.attrib['name']
                mech_chassis = mech.attrib['chassis']
                mechs[mech_id] = {
                    'name': mech_name
                }

    return mechs


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

                if 'InheritFrom' in weapon.attrib:
                    continue

                weapon_stats = weapon.find('WeaponStats')
                weapon_loc = weapon.find('Loc')

                weapon_id = weapon.attrib['id']
                weapon_name = localizations[weapon_loc.attrib['nameTag'][1:].lower()]
                weapon_type = weapon_stats.attrib['type']

                weapons[weapon_id] = {
                    # Weapon details
                    'id': weapon_id,
                    'name':  weapon_name,
                    'type': weapon_type,

                    # Equipment stats
                    'health': weapon_stats.attrib['Health'],
                    'tons': weapon_stats.attrib['tons'],
                    'slots': weapon_stats.attrib['slots'],

                    # Weapon stats
                    'damage': weapon_stats.attrib['damage'],
                    'heat': weapon_stats.attrib['heat'],
                    'duration': weapon_stats.attrib['duration'],
                    'cooldown': weapon_stats.attrib['cooldown'],

                    # Weapon behavior

                    'velocity': weapon_stats.attrib['speed'],
                    'count': weapon_stats.attrib['numFiring'],
                    'delay': weapon_stats.attrib['volleydelay'],

                    'minRange': weapon_stats.attrib['minRange'],
                    'optRange': weapon_stats.attrib['longRange'],
                    'maxRange': weapon_stats.attrib['maxRange'],

                    # Ghost heat
                    'penalty': weapon_stats.attrib['heatpenalty'] if 'heatpenalty' in weapon_stats.attrib else 0,
                    'penaltyLimit': weapon_stats.attrib['minheatpenaltylevel'] if 'minheatpenaltylevel' in weapon_stats.attrib else 0,
                    'penaltyGroup': weapon_stats.attrib['heatPenaltyID'] if 'heatPenaltyID' in weapon_stats.attrib else 0
                }

    return weapons

if __name__ == '__main__':
    # Extract data
    localizations = map_localizations()
    weapons = map_weapons(localizations)

    # Export JSON
    locale_json_file = 'out/locale.json'
    if not os.path.exists(os.path.dirname(locale_json_file)):
        try:
            os.makedirs(os.path.dirname(locale_json_file))
        except OSError as exc: # Guard against race condition
            if exc.errno != errno.EEXIST:
                raise

    with open(locale_json_file, 'w') as locale_json:
        locale_json.write(json.dumps(map_localizations(), indent=2))

    # Export weapon JSON
    weapon_json_file = 'out/weapons.json'
    if not os.path.exists(os.path.dirname(weapon_json_file)):
        try:
            os.makedirs(os.path.dirname(weapon_json_file))
        except OSError as exc: # Guard against race condition
            if exc.errno != errno.EEXIST:
                raise

    with open(weapon_json_file, 'w') as weapon_json:
        weapon_json.write(json.dumps(weapons, indent=2))
