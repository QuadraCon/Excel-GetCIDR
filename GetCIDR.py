import os
import pandas as pd
import ipaddress

def ValidIP(address):
    try:
        ipaddress.ip_address(address)
        return True
    except Exception as e:
        return False

def ip_to_bin(ip):
    # Convert an IP address to its binary representation
    ip_str = str(ip)  # Ensure the IP address is treated as a string
    return ''.join(format(int(octet), '08b') for octet in ip_str.split('.'))

def common_prefix_length(bin_net, bin_bcast):
    # Find the common prefix length of the network and broadcast binary strings
    return next((i for i in range(len(bin_net)) if bin_net[i] != bin_bcast[i]), len(bin_net))

def deduce_cidr(network_address, broadcast_address):
    # Deduce the CIDR notation from network and broadcast addresses
    if ValidIP(network_address) & ValidIP(broadcast_address):
        bin_net = ip_to_bin(network_address)
        bin_bcast = ip_to_bin(broadcast_address)
        prefix_length = common_prefix_length(bin_net, bin_bcast)
        return f"{network_address}/{prefix_length}"

def find_network_and_broadcast_addresses(root_directory):
    results = []
    temp_data = {}

    # Traverse through all folders and files
    for subdir, _, files in os.walk(root_directory):
        for file in files:
            # Check if the file is an Excel file
            if file.endswith('.xlsx') or file.endswith('.xls'):
                file_path = os.path.join(subdir, file)
                try:
                    # Load the Excel file
                    excel_file = pd.ExcelFile(file_path)
                    # Iterate through each sheet in the Excel file
                    for sheet_name in excel_file.sheet_names:
                        sheet = excel_file.parse(sheet_name)
                        # Ensure the sheet has at least two columns
                        if sheet.shape[1] >= 2:
                            for index, row in sheet.iterrows():
                                if pd.notna(row.iloc[1]):
                                    key = (file_path, sheet_name)
                                    if row.iloc[1].lower() == 'network address':
                                        network_address = row.iloc[0]
                                        if key not in temp_data:
                                            temp_data[key] = {}
                                        temp_data[key]['Network address'] = network_address
                                    elif row.iloc[1].lower() == 'broadcast address':
                                        broadcast_address = row.iloc[0]
                                        if key not in temp_data:
                                            temp_data[key] = {}
                                        temp_data[key]['Broadcast address'] = broadcast_address
                except Exception as e:
                    print(f"Failed to read {file_path}: {e}")

    for key, addresses in temp_data.items():
        if 'Network address' in addresses and 'Broadcast address' in addresses:
            network_address = addresses['Network address']
            broadcast_address = addresses['Broadcast address']
            cidr = deduce_cidr(network_address, broadcast_address)
            results.append({
                'CIDR': cidr
            })

    return results

# Replace with the path to your root directory
root_directory = r'C:\Path\To\Folder'
results = find_network_and_broadcast_addresses(root_directory)

# Save results to a CSV file
results_df = pd.DataFrame(results)
results_df.to_csv('network_and_broadcast_addresses_with_cidr.csv', index=False)

print(f"Results saved to network_and_broadcast_addresses_with_cidr.csv")