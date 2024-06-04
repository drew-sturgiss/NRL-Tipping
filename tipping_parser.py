import argparse
import datetime
import gdown
import pandas

def download_onedrive_file(local_file_location, onedrive_file_location):
    '''Downloads a specified onedrive file to a specified local path
    '''

def enrich_tipping_sheet(round_sheet, players_sheet):
    # Add a column to the round sheet to store the player's name
    round_sheet.insert(2, 'Name')
    # Add a column to the round sheet to store the player's rank
    round_sheet.insert(3, 'Rank')
    # Add a column to the round sheet to store the player's points


    # Add the players name
    ## For each row in the round sheet
    for each_tips_entry in round_sheet.iterrows():
        ## For each row in the players sheet
        for each_player in players_sheet.iterrows():
            ## If the player's ID matches the ID in the round sheet
            if each_tips_entry['Email Address'] == each_player['Email']:
                ## Add the player's name to the round sheet
                each_tips_entry['Name'] = each_player['Name']
                ## Exit the loop
                break

    # Sort by name
    round_sheet = round_sheet.sort_values(by='Name')
    
    # Add Excel formula columns]
    player = 1
    for each_tips_entry in round_sheet.iterrows():
        row_number = player + 1
        # Add rank forumla
        each_tips_entry['Rank'] = f"=VLOOKUP(B2, Players!A:D, 3, FALSE)"
        # Add points formula
        each_tips_entry['Points'] = f"=VLOOKUP(B2, Players!A:D, 4, FALSE)"

        # Increment the player number
        player += 1

    # Return the enriched round sheet
    return round_sheet

def get_round_number(footy_draw):
    '''Takes a dataframe containing the footy draw as an argument,
    and returns the round number of the next game after the current date
    '''
    # Create a variable representing the current time
    current_time = datetime.datetime.now()

    # Create a variable representing the format the draw sheet stores game times in
    game_date_format = '%-d/%-m/%Y %H:%M'

    # Create a variable representing the round number
    round_number = 0

    for each_game in footy_draw.iterrows():
        # Create a variable representing the time of the game
        game_time = datetime.strptime(each_game['Date'], game_date_format)
        # If the game occurs after right now
        if game_time > current_time:
            # Store the game's round number
            round_number = each_game['Round']
            # Exit the while loop
            break

    # Return the round number
    return round_number

def get_tipping_validity(round_sheet):
    ''' Check if the tipping sheet constitutes a valid tipping sheet
    '''
    tipping_sheet_validity = True
    if len(tipping_sheet_validity.index) != 6:
        tipping_sheet_validity = False
    
    return tipping_sheet_validity

def get_tips_sheet_url(round_number, weekly_forms):
    ''' Returns the URL for a tips entry spreadsheet for the round number
    specified by a given url
    '''
    # Identify the appropriate row given by it matching the round number
    round = weekly_forms.loc[weekly_forms['Week'] == f"Week {round_number}"]
    
    # Record the URL of that weeks tipping sheet
    sheet_url = round['Sheet Link']

    # Return the link to the sheet
    return sheet_url

def ingest_google_sheet_to_dataframe(google_sheet_url):
    ''' Ingest a google speadsheet at a given url into a dataframe, and
    return the dataframe
    '''

def ingest_onedrive_sheet_to_dataframe(onedrive_url):
    '''
    Ingest a onedrive spreadsheet at a given url into a dataframe, and
    return the dataframe
    '''

def process_tipping_entries(round_sheet, tips_sheet):
    ''' Take a given tipping entry and add it to the tipping sheet
    '''
    for each_tipping_entry in round_sheet.iterrows():
        # For every game column
        for each_tipping_entry_column in each_tipping_entry[5:]:
            # Construct the information for each tip
            ## Calculate the row number (required for Excel functions)
            row_number = len(tips_sheet.index) + 1
            ## Determine the name of the game by grabbing it from the column name
            game_column = each_tipping_entry_column
            ## Use a formula to let the Excel sheet look up the round number
            week_column = f"=VLOOKUP(A{row_number}, 'Footy Draw'!A:B, 2, FALSE)"
            ## Store the player's name
            player_column = each_tipping_entry['Name']
            ## Store the player's pick
            pick_column = each_tipping_entry[each_tipping_entry_column]
            ## Use a formula to let Excel lookup the winner
            winner_column = f"=VLOOKUP(A{row_number}, 'Footy Draw'!A:I, 8, FALSE)"
            ## Use a formula to let excel determine if the player gets a point
            points_column = f"=IF(EXACT(D{row_number}, E{row_number}), 1, 0)"

            # Add the information to the tipping sheet
            ## Construct a dictionary to store the information
            tipping_row = {'Game': game_column, 'Week': week_column, 'Player': player_column, 'Pick': pick_column, 'Winner': winner_column, 'Points': points_column}
            ## Add the new to the dataframe
            tips_sheet = tips_sheet.append(tipping_row, ignore_index=True)

    # Return the updated tipping sheet
    return tips_sheet

def main():
    # Download master spreadsheet
    tipping_spreadsheet = ingest_onedrive_sheet_to_dataframe()

    # Determine current round
    round_number = get_round_number(tipping_spreadsheet['Footy Draw'])

    # Download round entries sheet
    sheet_url = get_tips_sheet_url(round_number, tipping_spreadsheet['Weekly Forms'])
    round_sheet = ingest_google_sheet_to_dataframe(sheet_url)
    
    # Enrich the tipping sheet with tipper's name
    round_sheet = enrich_tipping_sheet(round_sheet, tipping_spreadsheet['Players'])

    # Confirm if the tipping sheet is valid
    if not get_tipping_validity(round_sheet):
        # Alert if not
        print("Tipping sheet is not valid")
        return

    # Record tipping entries in master record
    tipping_spreadsheet['Tips'] = process_tipping_entries(round_sheet, tipping_spreadsheet['Tips'])

    # Construct round summary sheet
    tipping_spreadsheet['Tipping Sheet'] = round_sheet

    # Email tipping sheet

    ## Generic functions
    # Email results

    # Format tipping sheet

    # Apply conditional formatting to tipping sheet column

    # Upload spreadsheet