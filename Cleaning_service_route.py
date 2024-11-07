from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import requests
import webbrowser

# Google Sheets API setup
credentials = Credentials.from_service_account_file("C:/Users/QN2505/Cleaning Servie Route/credentials.json")  # Replace with path to your credentials file
service = build('sheets', 'v4', credentials=credentials)

SPREADSHEET_ID = '1PurK5h1ewrsvKDtDhjdqLBFs8dM5Y4avS4i56L6gW0w'  # Replace with your actual Spreadsheet ID
RANGE_NAME = 'Sheet1!B2:B5'  # Adjust the range to your data location

API_KEY = 'AIzaSyA2JoLgfxQ6RfcmojL8xW_RSaNjWv4uC3Y'  # Replace with your actual Map Platform API key

def get_addresses_from_sheet():
    """Fetches addresses from a specified Google Sheets range."""
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    values = result.get('values', [])

    if not values:
        print("No data found.")
        return []

    # Extract addresses from each row
    addresses = [row[0] for row in values]
    return addresses

def get_directions(start_address, waypoints):
    """Fetches optimized directions from the Google Directions API."""
    waypoints_str = '|'.join(waypoints)
    url = f"https://maps.googleapis.com/maps/api/directions/json?origin={start_address}&destination={start_address}&waypoints=optimize:true|{waypoints_str}&key={API_KEY}"

    response = requests.get(url)
    directions = response.json()

    if directions['status'] == 'OK':
        return directions
    else:
        print(f"Error: {directions['status']}")
        return None

def main():
    """Main function to retrieve addresses, get directions, and display route."""
    addresses = get_addresses_from_sheet()  # Fetch addresses dynamically

    if len(addresses) < 2:
        print("At least two addresses are needed to calculate a route.")
        return

    start_address = addresses[0]
    waypoints = addresses[1:]

    # Get directions
    directions = get_directions(start_address, waypoints)

    # Print the optimized route instructions
    if directions:
        print("Optimized Route Instructions:")
        for leg in directions['routes'][0]['legs']:
            for step in leg['steps']:
                print(step['html_instructions'])

    # Construct and open the route in Google Maps
    waypoints_str = "|".join(waypoints)
    google_maps_url = (
        f"https://www.google.com/maps/dir/?api=1"
        f"&origin={start_address}"
        f"&destination={start_address}"
        f"&waypoints={waypoints_str}"
    )
    webbrowser.open(google_maps_url)

if __name__ == '__main__':
    main()
