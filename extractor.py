# -*- coding: utf-8 -*-
"""
extractor.py

Extracts MWO game data to a more accessible format.
"""

import os
import zipfile
import winreg
from lxml import etree as ET

# Ensure wdir is current
abs_path = os.path.abspath(__file__)
dir_name = os.path.dirname(abs_path)
os.chdir(dir_name)

from locations import gamedata_pak, weapon_file, mech_file, english_pak, localization_file,

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

if __name__ == '__main__':
    localizations = map_localizations()

    print(localizations['kgc-000bs'])
    print(localizations['ml_desc'])
    print(localizations['ac20_desc'])
    print(map_mechs())
