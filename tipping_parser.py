import argparse
import datetime
import gdown
import pandas

def determine_round_number(footy_draw):
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
        game_time = strptime(each_game['Date'], game_date_format)
        # If the game occurs after right now
        if game_time > current_time:
            # Store the game's round number
            round_number = each_game['Round']
            # Exit the while loop
            break

    # Return the round number
    return round_number


def download_onedrive_file(local_file_location, onedrive_file_location):
    '''Downloads a specified onedrive file to a specified local path
    '''

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
    ''' Ingest a google sheet at a given url into a dataframe, and
    return the dataframe
    '''

def process_tipping_entry(tips_entry):


def main():
    # Download master spreadsheet

    # Determine current round
    round_number = determine_round_number(tipping_spreadsheet['Footy Draw'])

    # Download round entries sheet
    sheet_url = get_tips_sheet_url(round_number), tipping_spreadsheet['Weekly Forms'])
    tipping_sheet = ingest_google_sheet_to_dataframe(sheet_url)

    # Confirm correct number of entries
        # Alert if not
        # End here

    # Add entries to spreadsheet

    # Construct round summary sheet

    # Email tipping sheet

    ## Generic functions
    # Email results

    # Format tipping sheet

    # Apply conditional formatting to tipping sheet column

    # Upload spreadsheet