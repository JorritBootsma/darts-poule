import glob
import os

import pandas as pd
from functools import reduce

def format_name(name):
    lastname = name.split(" ")[-2]
    first_three_last_name = lastname[:3]
    return first_three_last_name


def format_column_name(home_player, away_player):
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
        files.append(file)
    return files


def extract_person_from_path(path_):
    filename = os.path.basename(path_)
    person = filename.split(" ")[-1].split(".")[0]
    return person


paths = find_paths_from("/Users/jorrit/Documents/Prive/darts_poule/")

column_pairs_dict = {
    "Ronde 1 Links": ("2023 PDC WK DARTS", "Unnamed: 2"),  # Ronde 1
    "Ronde 2 Links": ("Unnamed: 3", "Unnamed: 4"),  # Ronde 2
    "Ronde 3 Links": ("Unnamed: 5", "Unnamed: 6"),  # Ronde 3
    "Achtste finale Links": ("Unnamed: 7", "Unnamed: 8"),  # Achtste finale
    "Kwart finale Links": ("Unnamed: 9", "Unnamed: 10"),  # Kwart finale
    "Halve finale Links": ("Unnamed: 11", "Unnamed: 12"),  # Halve finale
    "Halve finale Rechts": ("Unnamed: 20", "Unnamed: 19"),  # Halve finale
    "Kwart finale Rechts": ("Unnamed: 22", "Unnamed: 21"),  # Kwart finale
    "Achtste finale Rechts": ("Unnamed: 24", "Unnamed: 23"),  # Achtste finale
    "Ronde 3 Rechts": ("Unnamed: 26", "Unnamed: 25"),  # Ronde 3
    "Ronde 2 Rechts": ("Unnamed: 28", "Unnamed: 27"),  # Ronde 2
    "Ronde 1 Rechts": ("Unnamed: 30", "Unnamed: 29"),  # Ronde 1
}

# Loop over the different columns that contain the player names and the scores
for round_name, column_pair in column_pairs_dict.items():
    matches_column = column_pair[0]
    scores_column = column_pair[1]

    all_score_dfs = []

    # Loop over the different Excel files (one excel file per person)
    for path in paths:
        # First column contains the contender's name
        col_names = []
        person_name = extract_person_from_path(path)
        print(f"\n\n Contender: {person_name}")

        # Start each row with the name of the contender
        col_values = []
        # col_values.append(person_name)

        df = pd.read_excel(path)

        matches_ronde_1 = df[matches_column]
        scores_ronde_1 = df[scores_column]

        df = pd.DataFrame(
            {
                "matches": matches_ronde_1,
                "scores": scores_ronde_1
            }
        )
        filtered_df = df[df['scores'].notnull()]
        filtered_df = filtered_df[~filtered_df['scores'].astype(str).str.contains("Best")]

        i = 0
        home_players = []
        home_scores = []
        away_players = []
        away_scores = []
        for row in filtered_df.iterrows():
            if i%2 == 0:
                home_players.append(row[1].matches)
                home_scores.append(row[1].scores)
            else:
                away_players.append(row[1].matches)
                away_scores.append(row[1].scores)
            i += 1

        print(round_name)

        for i in range(len(home_players)):
            match_i = format_column_name(home_players[i], away_players[i])
            col_names.append(match_i)

        for i in range(len(home_players)):
            score_i = format_score(home_scores[i], away_scores[i])
            col_values.append(score_i)

        print(len(col_names), col_names)
        print(len(col_values), col_values)

        data_dict = {}
        for key, value in list(zip(col_names, col_values)):
            data_dict[key] = [value]

        contender_score_df = pd.DataFrame(
            data=data_dict,
            index=[person_name]
        )

        print(contender_score_df)

        all_score_dfs.append(contender_score_df)

    # scores_df = reduce(lambda df1, df2: pd.merge(df1, df2, on="Deelnemer"), all_score_dfs)
    scores_df = pd.concat(all_score_dfs, sort=False)
    print(scores_df)


    print()
    print()
    # scores_df = pd.DataFrame(all_scores, columns=all_col_names)
    # print(scores_df)

    with pd.ExcelWriter(f'Darts_scores_{round_name}.xlsx') as writer:
        scores_df.to_excel(writer, sheet_name=round_name, index=True)

