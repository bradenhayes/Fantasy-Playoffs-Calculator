import os
import sys
import requests
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
import urllib3
from string import ascii_uppercase

def get_column_letter(index):
    """Convert numeric index to spreadsheet column letter (A, B, C, ..., Z, AA, AB, etc.)"""
    result = ""
    while index >= 0:
        result = ascii_uppercase[index % 26] + result
        index = index // 26 - 1
    return result

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
        """Create mapping of player names to IDs, including only offensive players"""
        print("Creating player name to ID mapping (offensive players only)...")
        players = self.get_players_info()
        name_to_id = {}
        
        # Define offensive positions
        offensive_positions = {'QB', 'RB', 'WR', 'TE', 'K', 'FB'}
        
        for player_id, player_info in players.items():
            if player_info:
                position = player_info.get('position')
                # Only process players in offensive positions
                if position in offensive_positions:
                    full_name = f"{player_info.get('first_name', '')} {player_info.get('last_name', '')}".strip().lower()
                    if full_name:
                        name_to_id[full_name] = player_id
        
        print(f"Mapped {len(name_to_id)} offensive players")
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
        # Get all sheets
        sheets_data = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        all_sheets = sheets_data.get('sheets', [])
        
        # Filter for week score sheets and sort them
        week_sheets = []
        for sheet in all_sheets:
            title = sheet['properties']['title']
            for week_num in range(1, 5):
                if get_round_name(week_num) in title:
                    week_sheets.append((week_num, title))
        
        week_sheets.sort()
        
        if not week_sheets:
            print("No weekly score sheets found!")
            return
        
        # Initialize dictionaries
        totals = {}
        weekly_scores = {}
        
        # Process each week's sheet
        for week_num, sheet_title in week_sheets:
            print(f"Processing {sheet_title}...")
            
            # Get the sheet's dimensions first
            sheet_metadata = service.spreadsheets().get(
                spreadsheetId=spreadsheet_id,
                ranges=[f"'{sheet_title}'!A:Z"],
                fields='sheets(properties(gridProperties))'
            ).execute()
            
            grid_props = sheet_metadata['sheets'][0]['properties']['gridProperties']
            max_col = grid_props.get('columnCount', 26)  # Default to 26 if not found
            
            # Get the data using dynamic range
            range_name = f"'{sheet_title}'!A1:{get_column_letter(max_col-1)}{grid_props.get('rowCount', 100)}"
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
            
            # Get total row and update running totals
            total_row = values[-1][1:]  # Skip 'TOTAL' label
            for owner_idx, owner in enumerate(owners):
                try:
                    week_total = float(total_row[owner_idx])
                    totals[owner] += week_total
                    weekly_scores[owner].append(week_total)
                except (ValueError, IndexError):
                    print(f"Warning: Could not process total for {owner} in week {week_num}")
                    weekly_scores[owner].append(0.0)
        
        # Create or update Running Totals sheet
        sheet_title = "Running Totals"
        
        # Check if sheet exists
        sheet_exists = any(
            sheet['properties']['title'] == sheet_title 
            for sheet in sheets_data.get('sheets', [])
        )
        
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
        
        # Prepare values
        values = []
        values.append(['Round'] + list(totals.keys()))
        
        # Add weekly scores
        for week_idx, (week_num, _) in enumerate(week_sheets):
            round_name = get_round_name(week_num)
            week_row = [round_name]
            for owner in totals.keys():
                try:
                    week_row.append(weekly_scores[owner][week_idx])
                except IndexError:
                    week_row.append(0.0)
            values.append(week_row)
        
        # Add running totals
        values.append(['TOTAL'] + list(totals.values()))
        
        # Update sheet with dynamic range
        last_col = get_column_letter(len(totals.keys()))
        range_name = f"'{sheet_title}'!A1:{last_col}{len(values)}"
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
    credentials = service_account.Credentials.from_service_account_file(
        get_resource_path("key.json"), 
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    
    service = build("sheets", "v4", credentials=credentials)
    sleeper = SleeperFantasyAPI()
    
    try:
        # Get sheet dimensions first
        sheet_metadata = service.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            fields='sheets(properties(gridProperties))'
        ).execute()
        
        grid_props = sheet_metadata['sheets'][0]['properties']['gridProperties']
        max_col = grid_props.get('columnCount', 26)
        
        # Read roster with dynamic range
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"A1:{get_column_letter(max_col-1)}{grid_props.get('rowCount', 100)}"
        ).execute()
        
        values = result.get('values', [])
        if not values:
            print("No data found in spreadsheet")
            return
        
        headers = values[0][1:]  # Skip first column, get owner names
        positions = [row[0] for row in values[1:] if row]  # Get positions, skip empty rows
        
        # Store scores
        scores = {owner: {} for owner in headers}
        
        # Process rosters
        for owner_idx, owner in enumerate(headers):
            print(f"\nProcessing roster for {owner}...")
            for pos_idx, pos in enumerate(positions):
                try:
                    player_name = values[pos_idx + 1][owner_idx + 1]
                    points = sleeper.get_weekly_stats_for_player(player_name, week)
                    scores[owner][pos] = {
                        'player': player_name,
                        'points': points
                    }
                    print(f"{pos}: {player_name} - {points} points")
                except IndexError:
                    print(f"Warning: Missing data for {owner} at position {pos}")
                    scores[owner][pos] = {
                        'player': 'MISSING',
                        'points': 0.0
                    }
        
        # Create scores sheet
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
        
        # Prepare values
        values = [['Position'] + list(scores.keys())]
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
        
        # Update sheet with dynamic range
        last_col = get_column_letter(len(headers))
        range_name = f"'{sheet_title}'!A1:{last_col}{len(values)}"
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
    spreadsheet_id = ""  # Add your spreadsheet ID here
    
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
