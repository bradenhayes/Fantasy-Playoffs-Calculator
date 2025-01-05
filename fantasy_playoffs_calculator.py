import os
import sys
import requests
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
import urllib3

def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class SleeperFantasyAPI:
   def __init__(self):
       """Initialize the Sleeper API client"""
       self.base_url = "https://api.sleeper.app/v1"
       self.current_nfl_season = 2024
       self.headers = {
           'User-Agent': 'Mozilla/5.0',
           'Accept': 'application/json'
       }
       # Create player mapping when initialized
       self.player_map = self.create_player_name_map()
   
   def get_players_info(self) -> dict:
       """Get all players information"""
       endpoint = f"{self.base_url}/players/nfl"
       response = requests.get(
           endpoint,
           headers=self.headers,
           verify=False
       )
       response.raise_for_status()
       return response.json()

   def create_player_name_map(self) -> dict:
       """Create mapping of player names to IDs"""
       print("Creating player name to ID mapping...")
       players = self.get_players_info()
       name_to_id = {}
       for player_id, player_info in players.items():
           if player_info:
               full_name = f"{player_info.get('first_name', '')} {player_info.get('last_name', '')}".strip().lower()
               name_to_id[full_name] = player_id
       print(f"Mapped {len(name_to_id)} players")
       return name_to_id
   
   def get_player_stats(self, week: int) -> dict:
       """Fetch stats for a specific week using Sleeper's PPR scoring"""
       endpoint = f"{self.base_url}/stats/nfl/post/{self.current_nfl_season}/{week}"
       print(f"Fetching stats from: {endpoint}")
       
       response = requests.get(
           endpoint, 
           headers=self.headers,
           verify=False
       )
       
       print(f"Response status: {response.status_code}")
       
       if response.status_code != 200:
           print(f"Error content: {response.text}")
           response.raise_for_status()
           
       return response.json()
   
   def get_weekly_stats_for_player(self, player_name: str, week: int) -> float:
       """Get weekly PPR points for a specific player"""
       player_id = self.player_map.get(player_name.lower())
       if not player_id:
           print(f"WARNING: Could not find player ID for {player_name}")
           return 0.0
       
       stats = self.get_player_stats(week)
       player_stats = stats.get(player_id, {})
       
       print(f"\nLooking up stats for {player_name} (ID: {player_id})")
       print(f"Full stats received: {player_stats}")
       
       # Try to find fantasy points in the stats
       if 'pts_ppr' in player_stats:
           return float(player_stats['pts_ppr'])
       elif 'fantasy_points_ppr' in player_stats:
           return float(player_stats['fantasy_points_ppr'])
       elif 'fantasy_points' in player_stats:
           return float(player_stats['fantasy_points'])
       
       # If no points found, show what data we did get
       print(f"Available stat keys: {list(player_stats.keys())}")
       return 0.0

def get_round_name(week_num):
   """Convert week number to round name"""
   playoff_round_names = {
       1: "Round 1",
       2: "Round 2", 
       3: "Round 3",
       4: "Championship"
   }
   return playoff_round_names.get(week_num, f"Week {week_num}")

def update_totals_sheet(service, spreadsheet_id, current_week):
   """Update the running totals sheet by processing all weekly score sheets"""
   print("\nUpdating running totals...")
   
   try:
       # First, get a list of all sheets to find our weekly sheets
       sheets_data = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
       all_sheets = sheets_data.get('sheets', [])
       
       # Filter for week score sheets and sort them
       week_sheets = []
       for sheet in all_sheets:
           title = sheet['properties']['title']
           for week_num in range(1, 5):  # Weeks 1-4 for playoffs
               if get_round_name(week_num) in title:
                   week_sheets.append((week_num, title))
       
       week_sheets.sort()  # Sort by week number
       
       if not week_sheets:
           print("No weekly score sheets found!")
           return
       
       # Initialize totals dictionary
       totals = {}
       weekly_scores = {}
       
       # Process each week's sheet
       for week_num, sheet_title in week_sheets:
           print(f"Processing {sheet_title}...")
           
           # Get the data from this sheet
           range_name = f"'{sheet_title}'!A1:Z100"  # Adjust range as needed
           result = service.spreadsheets().values().get(
               spreadsheetId=spreadsheet_id,
               range=range_name
           ).execute()
           
           values = result.get('values', [])
           if not values:
               continue
               
           # Get owner names from first row (skip 'Position' column)
           owners = values[0][1:]
           
           # Initialize dictionaries if this is the first sheet
           if not totals:
               totals = {owner: 0.0 for owner in owners}
               weekly_scores = {owner: [] for owner in owners}
           
           # Get total row (last row) and update running totals
           total_row = values[-1][1:]  # Skip 'TOTAL' label
           for owner_idx, owner in enumerate(owners):
               try:
                   week_total = float(total_row[owner_idx])
                   totals[owner] += week_total
                   weekly_scores[owner].append(week_total)
               except (ValueError, IndexError):
                   print(f"Warning: Could not process total for {owner} in week {week_num}")
                   weekly_scores[owner].append(0.0)
       
       # Check if Running Totals sheet exists and create if it doesn't
       sheet_title = "Running Totals"
       sheet_exists = False
       
       # Check if Running Totals sheet already exists
       sheets_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
       for sheet in sheets_metadata.get('sheets', ''):
           if sheet['properties']['title'] == sheet_title:
               sheet_exists = True
               break
       
       # Create Running Totals sheet if it doesn't exist
       if not sheet_exists:
           body = {
               'requests': [{
                   'addSheet': {
                       'properties': {
                           'title': sheet_title
                       }
                   }
               }]
           }
           service.spreadsheets().batchUpdate(
               spreadsheetId=spreadsheet_id,
               body=body
           ).execute()
       
       # Prepare values for the totals sheet
       values = []
       
       # Header row with owner names
       values.append(['Round'] + list(totals.keys()))
       
       # Add weekly scores with round names
       for week_idx, (week_num, _) in enumerate(week_sheets):
           round_name = get_round_name(week_num)
           week_row = [round_name]
           for owner in totals.keys():
               try:
                   week_row.append(weekly_scores[owner][week_idx])
               except IndexError:
                   week_row.append(0.0)
           values.append(week_row)
       
       # Add running totals row
       values.append(['TOTAL'] + list(totals.values()))
       
       # Update the sheet
       range_name = f"'{sheet_title}'!A1:{chr(65 + len(totals))}{len(values)}"
       body = {
           'values': values
       }
       service.spreadsheets().values().update(
           spreadsheetId=spreadsheet_id,
           range=range_name,
           valueInputOption='RAW',
           body=body
       ).execute()
       
       print(f"Successfully updated running totals!")
       print("\nCurrent Standings:")
       sorted_totals = sorted(totals.items(), key=lambda x: x[1], reverse=True)
       for owner, total in sorted_totals:
           print(f"{owner}: {total:.2f} points")
           
   except Exception as e:
       print(f"An error occurred updating totals: {e}")

def process_spreadsheet(spreadsheet_id: str, week: int):
   """Process the spreadsheet and update scores"""
   # Set up Google Sheets credentials
   credentials = service_account.Credentials.from_service_account_file(
       get_resource_path("key.json"), 
       scopes=["https://www.googleapis.com/auth/spreadsheets"]
   )
   
   service = build("sheets", "v4", credentials=credentials)
   
   # Initialize Sleeper API
   sleeper = SleeperFantasyAPI()
   
   try:
       # Read the roster sheet
       result = service.spreadsheets().values().get(
           spreadsheetId=spreadsheet_id,
           range="A1:F10"  # Adjust range as needed
       ).execute()
       
       values = result.get('values', [])
       if not values:
           print("No data found in spreadsheet")
           return
       
       # Process the roster data
       headers = values[0][1:]  # Skip first column, get owner names
       positions = [row[0] for row in values[1:]]  # Get positions
       
       # Store scores for each owner
       scores = {owner: {} for owner in headers}
       
       # Process each owner's roster
       for owner_idx, owner in enumerate(headers):
           print(f"\nProcessing roster for {owner}...")
           for pos_idx, pos in enumerate(positions):
               player_name = values[pos_idx + 1][owner_idx + 1]
               points = sleeper.get_weekly_stats_for_player(player_name, week)
               scores[owner][pos] = {
                   'player': player_name,
                   'points': points
               }
               print(f"{pos}: {player_name} - {points} points")
       
       # Create new sheet for scores using round name
       sheet_title = f"{get_round_name(week)} Scores"
       try:
           body = {
               'requests': [{
                   'addSheet': {
                       'properties': {
                           'title': sheet_title
                       }
                   }
               }]
           }
           service.spreadsheets().batchUpdate(
               spreadsheetId=spreadsheet_id,
               body=body
           ).execute()
       except Exception as e:
           print(f"Sheet might already exist: {e}")
       
       # Prepare values for the sheet
       values = [['Position'] + list(scores.keys())]  # Header row
       for pos in positions:
           row = [pos]
           for owner in scores:
               row.append(scores[owner][pos]['points'])
           values.append(row)
       
       # Add total row
       total_row = ['TOTAL']
       for owner in scores:
           total = sum(scores[owner][pos]['points'] for pos in positions)
           total_row.append(total)
       values.append(total_row)
       
       # Update the sheet
       range_name = f"'{sheet_title}'!A1:{chr(65 + len(headers))}{len(values)}"
       body = {
           'values': values
       }
       service.spreadsheets().values().update(
           spreadsheetId=spreadsheet_id,
           range=range_name,
           valueInputOption='RAW',
           body=body
       ).execute()
       
       print(f"\nSuccessfully updated {sheet_title}!")
       
       # Update running totals
       update_totals_sheet(service, spreadsheet_id, week)
       
   except Exception as e:
       print(f"An error occurred: {e}")

def get_actual_week(playoff_round):
   """Convert playoff round to actual week number"""
   return playoff_round  # No longer need to add 18

def get_user_input():
   """Get playoff round from user"""
   while True:
       try:
           print("\nFantasy Football Playoff Calculator")
           print("==================================")
           print("1: First Round")
           print("2: Second Round")
           print("3: Third Round")
           print("4: Championship")
           print("5: Exit")
           
           choice = input("\nEnter playoff round (1-5): ").strip()
           
           if choice == '5':
               print("Exiting program...")
               sys.exit(0)
           
           round_num = int(choice)
           if 1 <= round_num <= 4:
               return round_num
           else:
               print("Please enter a valid round number (1-4)")
       except ValueError:
           print("Please enter a valid number")

def main():
   # Suppress SSL verification warnings
   urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
   
   # Google Sheets spreadsheet ID
   spreadsheet_id = ""
   
   try:
       # Get playoff round from user
       playoff_round = get_user_input()
       
       # Calculate actual week number
       actual_week = get_actual_week(playoff_round)
       
       print(f"\nProcessing {get_round_name(actual_week)}")
       
       # Process the week
       process_spreadsheet(spreadsheet_id, actual_week)
       
       print("\nProcessing complete!")
       input("\nPress Enter to exit...")
       
   except Exception as e:
       print(f"\nAn error occurred: {e}")
       input("\nPress Enter to exit...")

if __name__ == "__main__":
   main()
