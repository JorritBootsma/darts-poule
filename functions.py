import glob
import os


def format_name(name):
    lastname = name.split(" ")[-2]
    first_three_last_name = lastname[:3]
    return first_three_last_name


def format_match_up(home_player, away_player):
    home_name = format_name(home_player)
    away_name = format_name(away_player)
    column_name = f"{home_name}-{away_name}"
    return column_name


def format_score(home_score, away_score):
    column_value = f"{int(home_score)} - {int(away_score)}"
    return column_value


def find_paths_from(folder):
    files = []
    for file in glob.glob(os.path.join(folder, "PDC*.xlsx")):
        # Skip the template file
        if "[NAAM]" in file:
            continue
        else:
            files.append(file)
    return files


def extract_person_from_path(path_):
    filename = os.path.basename(path_)
    person = filename.split(" ")[-1].split(".")[0]
    return person
