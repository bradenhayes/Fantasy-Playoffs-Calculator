import requests
import pandas as pd
from datetime import datetime

class SleeperFantasyAPI:
    def __init__(self):
        """Initialize the Sleeper API client"""
        self.base_url = "https://api.sleeper.app/v1"
        self.current_nfl_season = 2024
        self.headers = {
            'User-Agent': 'Mozilla/5.0',
            'Accept': 'application/json'
        }
    
    def get_player_stats(self, week: int) -> dict:
        """Fetch stats for a specific week using Sleeper's PPR scoring"""
        endpoint = f"{self.base_url}/stats/nfl/{self.current_nfl_season}/{week}"
        print(f"Fetching data from: {endpoint}")  # Debug print
        
        response = requests.get(
            endpoint, 
            headers=self.headers,
            verify=False
        )
        
        # Print response status and content for debugging
        print(f"Response status: {response.status_code}")
        
        if response.status_code != 200:
            print(f"Error content: {response.text}")
            response.raise_for_status()
            
        return response.json()
    
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
    
    def get_weekly_ppr_points(self, week: int) -> pd.DataFrame:
        """Get PPR points for all players in a specific week"""
        stats = self.get_player_stats(week)
        players = self.get_players_info()
        
        players_data = []
        for player_id, player_stats in stats.items():
            if player_stats and player_id in players:
                player_info = players[player_id]
                players_data.append({
                    'player_id': player_id,
                    'name': f"{player_info.get('first_name', '')} {player_info.get('last_name', '')}",
                    'position': player_info.get('position'),
                    'team': player_info.get('team'),
                    'ppr_points': player_stats.get('pts_ppr', 0),
                    'week': week
                })
        
        return pd.DataFrame(players_data)

def main():
    # Suppress SSL verification warnings
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    
    # Initialize the API
    sleeper = SleeperFantasyAPI()
    
    # Try multiple recent weeks in case some aren't available
    weeks_to_try = [13, 12, 11, 10]
    
    for week in weeks_to_try:
        try:
            print(f"\nAttempting to fetch Week {week} data...")
            week_points = sleeper.get_weekly_ppr_points(week)
            
            print(f"\nTop Week {week} Performers:")
            top_weekly = week_points.nlargest(10, 'ppr_points')[['name', 'position', 'team', 'ppr_points']]
            print(top_weekly.to_string(index=False))
            
            # Save to CSV
            week_points.to_csv(f'week_{week}_stats.csv', index=False)
            print(f"\nStats saved to week_{week}_stats.csv")
            
            # If we successfully get data, break the loop
            break
            
        except requests.exceptions.RequestException as e:
            print(f"Error fetching Week {week} data: {e}")
            continue
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            continue
    else:
        print("\nCouldn't fetch data for any of the attempted weeks.")

if __name__ == "__main__":
    main()
