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

character_data_to_sheet_fields_map = {
    'name': (0, 68),
    'background': (0, 72),
    'class': (0, 69),
    'subclass': None,
    'species': (0, 71),
    'level': (0, 69),
    'xp': (0, 73),
    'armor_class': (0, 140),
    'initiative': (0, 139),
    'speed': (0, 142),
    'hit_points_max': (0, 143),
    'hit_dice': (0, 145),
    'proficiency_bonus': (0, 141),
    'size': (3,23),
    'passive_perception': (0, 135),
    'passive_insight': (0, 136),
    'passive_investigation': (0, 137),
    'strength': (0, 74),
    'strength_mod': (0, 75),
    'dexterity': (0, 76),
    'dexterity_mod': (0, 77),
    'constitution': (0, 78),
    'constitution_mod': (0, 79),
    'intelligence': (0, 80),
    'intelligence_mod': (0, 81),
    'wisdom': (0, 82),
    'wisdom_mod': (0, 83),
    'charisma': (0, 84),
    'charisma_mod': (0, 85),
    'strength_saving_throw': (0, 86),
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
}

anchor_points_map = {
    "=== ARMOR ===": (None, None),
    "Resistances": (None, None)
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
        if ix:
            mapped_data[key] = dnd_beyond_data[ix[0]][ix[1]].strip()
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
    # abilities = ['strength', 'dexterity', 'constitution', 'intelligence', 'wisdom', 'charisma']
    # ability_saves_start = character_data_to_sheet_fields_map['strength_saving_throw'][1]
    # ability_saves_end = anchor_points_map['Resistances'][1]
    # for ability_index in character_data[0][ability_saves_start:ability_saves_end]:
    #     print("Ability:", ability_index)
    #     ability = abilities.pop(0)
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
        'credentials.json',
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
    parser.add_argument('sheet_name', nargs='?', default='Scratch 5E Character Sheet 2024',
                        help='Name of the Google spreadsheet to open')
    parser.add_argument("path", nargs="?", default="character_export.pdf", help="Path to PDF file")
    args = parser.parse_args()

    # Load character data from the specified PDF file
    character_data = load_character_data(args.path)
    find_anchor_points(character_data)

    # Print a minimal summary so the user can verify loading worked
    mapped_data = map_dnd_beyond_fields(character_data)
    mapped_data = cleanup_special_cases(character_data, mapped_data)
    print("Data", mapped_data)

    # Write the mapped data to the Google Sheet
    write_to_sheet(args.sheet_name, "Page 1", mapped_data)


if __name__ == '__main__':
    main()
