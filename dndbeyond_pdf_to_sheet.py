from __future__ import annotations
import argparse
from typing import Any, Dict
import pymupdf
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Define the scope for Google Sheets and Drive API
scope = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

anchor_points_map = {
    "ABILITY SAVE DC": (None, None),
    "=== ARMOR ===": (None, None),
    "Resistances": (None, None),
    "PERSONALITY TRAITS": (None, None),
}

# Each field is defined as a tuple of (page_index, anchor_point, line_index) where
# anchor_point is a string to search for and line_index is the line offset from that anchor point.
character_data_to_sheet_fields_map = {
    'name': (0, "ABILITY SAVE DC", 1),
    'class': (0, "ABILITY SAVE DC", 2),
    'level': (0, "ABILITY SAVE DC", 2),
    'subclass': None,
    'species': (0, "ABILITY SAVE DC", 4),
    'background': (0, "ABILITY SAVE DC", 5),
    'xp': (0, "ABILITY SAVE DC", 6),
    
    'initiative': (0, "=== ARMOR ===", -7),
    'armor_class': (0, "=== ARMOR ===", -6),
    'speed': (0, "=== ARMOR ===", -4),
    'hit_points_max': (0, "=== ARMOR ===", -3),
    'hit_dice': (0, "=== ARMOR ===", -1),
    'proficiency_bonus': (0, "=== ARMOR ===", -5),
    'size': (3, "PERSONALITY TRAITS", 3),

    'passive_perception': (0, "=== ARMOR ===", -11),
    'passive_insight': (0, "=== ARMOR ===", -10),
    'passive_investigation': (0, "=== ARMOR ===", -9),

    'strength': (0, "ABILITY SAVE DC", 7),
    'strength_mod': (0, "ABILITY SAVE DC", 8),
    'dexterity': (0, "ABILITY SAVE DC", 9),
    'dexterity_mod': (0, "ABILITY SAVE DC", 10),
    'constitution': (0, "ABILITY SAVE DC", 11),
    'constitution_mod': (0, "ABILITY SAVE DC", 12),
    'intelligence': (0, "ABILITY SAVE DC", 13),
    'intelligence_mod': (0, "ABILITY SAVE DC", 14),
    'wisdom': (0, "ABILITY SAVE DC", 15),
    'wisdom_mod': (0, "ABILITY SAVE DC", 16),
    'charisma': (0, "ABILITY SAVE DC", 17),
    'charisma_mod': (0, "ABILITY SAVE DC", 18),

    'strength_saving_throw': (0, "ABILITY SAVE DC", 19),
    'dexterity_saving_throw': None,
    'constitution_saving_throw': None,
    'intelligence_saving_throw': None,
    'wisdom_saving_throw': None,
    'charisma_saving_throw': None,
    'strength_proficiency': None,
    'dexterity_proficiency': None,
    'constitution_proficiency': None,
    'intelligence_proficiency': None,
    'wisdom_proficiency': None,
    'charisma_proficiency': None,
}

sheet_fields_to_sheet_cells_map = {
    'name':'D3',
    'background':'D9',
    'class':'V9',
    'subclass':'V13',
    'species':'D13',
    'level':'AO4',
    'xp':'AO11',
    'armor_class': 'AV6',
    'initiative': 'AN27',
    'speed': 'BA27',
    'hit_points_max': 'BP12',
    'hit_dice': 'BY12',
    'proficiency_bonus': 'D27',
    'size': 'BN27',
    'passive_abilities': 'CA27',
    'strength': 'M40',
    'strength_mod': 'D39',
    'dexterity': 'M61',
    'dexterity_mod': 'D60',
    'constitution': 'M86',
    'constitution_mod': 'D85',
    'intelligence': 'AE29',
    'intelligence_mod': 'V28',
    'wisdom': 'AE58',
    'wisdom_mod': 'V57',
    'charisma': 'AE87',
    'charisma_mod': 'V86',
    'strength_saving_throw': 'F49',
    'dexterity_saving_throw': 'F70',
    'constitution_saving_throw': 'F95',
    'intelligence_saving_throw': 'X38',
    'wisdom_saving_throw': 'X67',
    'charisma_saving_throw': 'X96',
    'strength_proficiency': 'D49',
    'dexterity_proficiency': 'D70',
    'constitution_proficiency': 'D95',
    'intelligence_proficiency': 'V38',
    'wisdom_proficiency': 'V67',
    'charisma_proficiency': 'V96',
}


def find_index_of_string_in_character_data(target: str, character_data: list) -> tuple:
    """
    Find the index of a target string in the character data.
    :param target: Target string to find
    :type target: str
    :param character_data: Character data to search
    :type character_data: list
    :return: Tuple of (page_index, line_index) if found, else None
    :rtype: tuple | None
    """
    for page_index, page in enumerate(character_data):
        for line_index, line in enumerate(page):
            if target in line:
                return (page_index, line_index)
    return None

def find_anchor_points(character_data: list) -> None:
    """
    Find anchor points in the character data.
    :param character_data: Character data to search
    :type character_data: list
    :return: None
    :rtype: None
    """
    for anchor, _ in anchor_points_map.items():
        index = find_index_of_string_in_character_data(anchor, character_data)
        anchor_points_map[anchor] = index

def load_character_data(path: str) -> list:
    """
    Load character data from a PDF file.
    :param path: Path to the PDF file
    :type path: str
    :return: List of text lines from each page
    :rtype: list
    """
    doc = pymupdf.open(path) # open a document
    text_list = []
    for page in doc: # iterate the document pages
        text = page.get_text()
        text_list.append(text.splitlines())
    return text_list

def map_dnd_beyond_fields(dnd_beyond_data: list) -> Dict[str, Any]:
    """Map D&D Beyond character data fields to another format.
    :param dnd_beyond_data: D&D Beyond character data
    :type dnd_beyond_data: Dict[str, Any]
    :return: Mapped character data
    :rtype: Dict[str, Any]
    """
    mapped_data = {}
    for key, ix in character_data_to_sheet_fields_map.items():
        print("Mapping field:", key, "Index:", ix)
        if ix is None:
            continue
        anchor = ix[1]
        relative_index = ix[2]
        absolute_index = anchor_points_map[anchor][1] + relative_index
        mapped_data[key] = dnd_beyond_data[ix[0]][absolute_index].strip()
    return mapped_data

def cleanup_special_cases(character_data: list, mapped_data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Handle special cases in the mapped data.
    :param character_data: Character data to search
    :type character_data: list
    :param mapped_data: Mapped character data
    :type mapped_data: Dict[str, Any]
    :return: Cleaned up character data
    :rtype: Dict[str, Any]
    """
    mapped_data['subclass'] = ''
    mapped_data['class'] = mapped_data.get('class').split(' ')[0]
    mapped_data['level'] = mapped_data.get('level').split(' ')[-1]
    mapped_data['passive_abilities'] = f"PER:{mapped_data.get('passive_perception')}, INS:{mapped_data.get('passive_insight')}, INV:{mapped_data.get('passive_investigation')}"
    
    # Handle ability saves
    abilities = ['strength', 'dexterity', 'constitution', 'intelligence', 'wisdom', 'charisma']
    anchor = character_data_to_sheet_fields_map['strength_saving_throw'][1]
    anchor_index = anchor_points_map[anchor][1]
    ability_saves_start = anchor_index + character_data_to_sheet_fields_map['strength_saving_throw'][2]
    # Initialize proficiency flags
    for ability in abilities:
        mapped_data[f"{ability}_proficiency"] = False
    current_ability = None
    ability_saves_end = anchor_points_map['Resistances'][1]
    # Iterate through ability saves
    for ability_index in range(ability_saves_start, ability_saves_end):
        if character_data[0][ability_index].strip() == '•':
            current_ability = abilities.pop(0)
            mapped_data[f"{current_ability}_proficiency"] = True
            continue
        elif not current_ability:
            current_ability = abilities.pop(0)
        mapped_data[f"{current_ability}_saving_throw"] = character_data[0][ability_index].strip()
        current_ability = None

    return mapped_data


def write_to_sheet(sheet_name: str, worksheet_name: str, mapped_data: Dict[str, Any]) -> None:
    """
    Write mapped character data to a Google Sheet.
    :param sheet_name: Name of the Google Spreadsheet
    :type sheet_name: str
    :param worksheet_name: Name of the worksheet/tab
    :type worksheet_name: str
    :param mapped_data: Mapped character data
    :type mapped_data: Dict[str, Any]
    :return: None
    :rtype: None
    """
    # Open the Google Sheet using the provided name
    # Authenticate with credentials and create a client to interact with Google Sheets
    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        'google-credentials.json',
        scope
        )
    client = gspread.authorize(credentials)
    sheet = client.open(sheet_name)
    ws = sheet.worksheet(worksheet_name)
    for field, cell in sheet_fields_to_sheet_cells_map.items():
        if field in mapped_data:
            ws.update_acell(cell, mapped_data[field])

def main() -> None:
    """
    Update and read a Google Sheet with character data from a JSON file.
    :return: None
    :rtype: None
    """

    # Parse command-line arguments
    parser = argparse.ArgumentParser(description='Update and read a Google Sheet')
    parser.add_argument('google_sheet_name', nargs='?', default='Scratch 5E Character Sheet 2024',
                        help='Name of the Google spreadsheet to open')
    parser.add_argument("character_sheet", nargs="?", default="character_export.pdf", help="Path to PDF file")
    args = parser.parse_args()

    # Load character data from the specified PDF file
    character_data = load_character_data(args.character_sheet)
    find_anchor_points(character_data)

    # Print a minimal summary so the user can verify loading worked
    mapped_data = map_dnd_beyond_fields(character_data)
    mapped_data = cleanup_special_cases(character_data, mapped_data)
    print("Data", mapped_data)

    # Write the mapped data to the Google Sheet
    write_to_sheet(args.google_sheet_name, "Page 1", mapped_data)

if __name__ == '__main__':
    main()
