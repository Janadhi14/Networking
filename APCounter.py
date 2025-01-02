import pandas as pd
from netmiko import ConnectHandler
import base64
import _secrets as _secrets
import re

# Define the getSwitchNames function to get input from the user for switches
def getSwitchNames():
    input_string = ''
    print("Enter the hostnames, each on a new line. Press Enter on an empty line to finish and press enter twice once you are done:")
    while True:
        line = input()
        if line.strip() == '':  # Exit the loop when the user presses Enter on an empty line
            break
        input_string += line + '\n'  # Add each line to the input string

    # Split the input into lines
    codes = input_string.strip().splitlines()

    # Create an array with each code wrapped in quotes
    formatted_codes = [f"{code}" for code in codes]
    return formatted_codes

# Get the switch hostnames from user input
switch_hostnames = getSwitchNames()

# Generate switches list dynamically
switches = []
for hostname in switch_hostnames:
    switches.append({
        'device_type': 'cisco_ios',  # You can change the device type if needed
        'host': hostname,
        'username': _secrets.username,
        'password': base64.b64decode(_secrets.password_b64.encode()).decode(),
        'port': 22,
        'secret': base64.b64decode(_secrets.enable_secret.encode()).decode()
    })

# Variable to accumulate total AP count across all switches
total_aps = 0
ap_model_counts = {}  # Dictionary to track counts of each AP model across all switches

# Initialize an empty list to store data for DataFrame
data_for_dataframe = []

# Function to handle connection and command execution
def execute_command_on_switch(switch):
    global total_aps, ap_model_counts  # Reference the global variables
    aps_per_switch = {}  # Dictionary to store the count of APs for the current switch
    switch_model = None  # Variable to store the switch model number
    switch_name = switch['host']  # Switch name (hostname)
    power_available = 0.0  # This will store the power available for the switch

    try:
        # Establish the SSH connection
        net_connect = ConnectHandler(**switch)

        # Get the switch model number from the 'show version' command
        version_resp = net_connect.send_command("show version")
        # Use regular expression to find the model number pattern
        model_match = re.search(r"Model Number\s+:\s+([^\s]+)", version_resp)
        if model_match:
            switch_model = model_match.group(1)
        else:
            # In case the model number isn't found, use the first matching string
            for line in version_resp.splitlines():
                if "Model" in line or "Cisco" in line:  # Adjust for your environment
                    switch_model = line.strip()
                    break
        
        # Clean up the model string by removing "Cisco IOS Software,"
        if switch_model and "Cisco IOS Software" in switch_model:
            switch_model = switch_model.replace("Cisco IOS Software,", "").strip()

        # Send command 'sh power inline' and capture output (use Genie or TextFSM)
        resp = net_connect.send_command("sh power inline", use_genie=True)

        # Genie response should be a dictionary-like structure; you need to check how this is returned
        if isinstance(resp, dict) and 'watts' in resp:
            # If Genie parsed it into a dict, process each interface for power information
            for k, v in resp['watts'].items():
                power_available += v['remaining']  # Sum the available power across all interfaces

        # If Genie fails or we have a raw response, handle manually (fallback)
        else:
            print(f"Raw response from {switch['host']}:\n{resp}")

        # Send command 'sh power inline' and capture output for APs (same as before)
        resp = net_connect.send_command("sh power inline", use_genie=True)

        # Process APs in the response
        for k, v in resp.get('interface', {}).items():
            device = v.get("device")
            # If device is a string and contains '-Z', it's assumed to be an AP
            if isinstance(device, str) and "-Z" in device:
                if device not in aps_per_switch:
                    aps_per_switch[device] = 1
                else:
                    aps_per_switch[device] += 1

        # Update total AP count across all switches
        total_aps += sum(aps_per_switch.values())

        # Update the AP count for each model across all switches (dynamically adding new models)
        for model, count in aps_per_switch.items():
            if model not in ap_model_counts:
                ap_model_counts[model] = count
            else:
                ap_model_counts[model] += count

        # Collect data for DataFrame
        row_data = {
            'Switch Name': switch_name,
            'Model': switch_model,
            'Power Avail.(Watts)': power_available,
            'Total APs in a Cabinet': sum(aps_per_switch.values())
        }

        # Add a column for each AP model dynamically
        for model in aps_per_switch:
            row_data[model] = aps_per_switch.get(model, 0)

        # Append row to the dataframe data list
        data_for_dataframe.append(row_data)

        # Print AP count per switch
        print(f"{switch['host']} (Model: {switch_model}): {aps_per_switch} - Power Available: {power_available}W")

        # Closing the connection after each switch is processed
        net_connect.disconnect()

    except Exception as e:
        print(f"Error connecting to {switch['host']}: {str(e)}")

# Iterate over all switches and execute the command
for switch in switches:
    execute_command_on_switch(switch)

# Print the total AP count after processing all switches
print(f"\nTotal AP count across all switches: {total_aps}")
print(f"Total AP count for each type of AP on all switches:")

# Print the count of each AP model across all switches
for model, count in ap_model_counts.items():
    print(f"{model}: {count}")


# Create a DataFrame from the collected data
df = pd.DataFrame(data_for_dataframe)

# Dynamically add the Total APs in a Cabinet column
df['Total APs in a Cabinet'] = df.apply(lambda row: sum(row[model] for model in ap_model_counts if model in row), axis=1)

# Replace NaN values with 0
df = df.fillna(0)

# Export the DataFrame to an Excel file
excel_filename = "Robertson.xlsx"
df.to_excel(excel_filename, index=False, engine='openpyxl')

print(f"Data has been exported to {excel_filename}")
