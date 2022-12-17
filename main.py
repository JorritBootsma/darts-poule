import pandas as pd

from functions import (
    extract_person_from_path,
    find_paths_from,
    format_score,
    format_match_up,
)

# Configuration variables
PATHS = find_paths_from("/Users/jorrit/Documents/Prive/darts_poule/")
COLUMN_PAIRS_DICT = {
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
for round_name, column_pair in COLUMN_PAIRS_DICT.items():
    # Extract the column names in which we can find the match-ups and the scores from
    # the config dictionary
    matches_column = column_pair[0]
    scores_column = column_pair[1]

    all_score_dfs = []

    # Loop over the different Excel files (one Excel file per person contending)
    for path in PATHS:
        person_name = extract_person_from_path(path)
        print(f"\n\n Contender: {person_name}")

        df = pd.read_excel(path)

        matches_this_round = df[matches_column]
        scores_this_round = df[scores_column]

        # Create dataframe, filter empty rows & irrelevant rows
        df = pd.DataFrame({"matches": matches_this_round, "scores": scores_this_round})
        filtered_df = df[df["scores"].notnull()]
        filtered_df = filtered_df[
            ~filtered_df["scores"].astype(str).str.contains("Best")
        ]

        # Extract match-ups using home- and away-players & extract corresponding scores
        i = 0
        home_players = []
        home_scores = []
        away_players = []
        away_scores = []
        for row in filtered_df.iterrows():
            if i % 2 == 0:
                home_players.append(row[1].matches)
                home_scores.append(row[1].scores)
            else:
                away_players.append(row[1].matches)
                away_scores.append(row[1].scores)
            i += 1

        print(round_name)

        # Format the column names (match-ups) and the values (scores)
        col_names = []
        col_values = []
        for i in range(len(home_players)):
            match_i = format_match_up(home_players[i], away_players[i])
            col_names.append(match_i)
            score_i = format_score(home_scores[i], away_scores[i])
            col_values.append(score_i)

        print(len(col_names), col_names)
        print(len(col_values), col_values)

        # Format as dictionary to load into Pandas DataFrame
        data_dict = {}
        for key, value in list(zip(col_names, col_values)):
            data_dict[key] = [value]

        contender_score_df = pd.DataFrame(data=data_dict, index=[person_name])

        # Make list of individual contender's dataframes
        all_score_dfs.append(contender_score_df)

    # Concatenate the different dataframes
    scores_df = pd.concat(all_score_dfs, sort=False)
    print(scores_df)

    # Write Excel file to disk for each round
    with pd.ExcelWriter(f"Darts_scores_{round_name}.xlsx") as writer:
        scores_df.to_excel(writer, sheet_name=round_name, index=True)
