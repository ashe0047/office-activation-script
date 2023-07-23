import argparse
import subprocess
import os
import re
import requests
from configparser import ConfigParser
import traceback

class OSPPModule:
    def __init__(self):
        self.ospp_commands = {
            # Global Options
            "activate_office_product": "/act",
            "install_product_key": "/inpkey:value",
            "uninstall_product_key": "/unpkey:value",
            "install_license": "/inslic:value",
            "display_license_info": "/dstatus",
            "display_all_license_info": "/dstatusall",
            "display_activation_failure_history": "/dhistoryacterr",
            "display_installation_id": "/dinstid",
            "activate_with_confirmation_id": "/actcid:value",
            "reset_licensing_status": "/rearm",
            "reset_license_status_with_skuid": "/rearm:value",
            "display_error_description": "/ddescr:value",

            # KMS Client Options
            "display_kms_activation_history": "/dhistorykms",
            "display_kms_client_machine_id": "/dcmid",
            "set_kms_host_name": "/sethst:value",
            "set_kms_port": "/setprt:value",
            "remove_kms_host_name": "/remhst",
            "permit_or_deny_kms_host_caching": "/cachst:value",
            "set_volume_activation_type": "/actype:value",
            "set_kms_srv_records_domain": "/skms-domain:value",
            "clear_kms_srv_records_domain": "/ckms-domain",

            # Token Options
            "display_installed_token_activation_issuance_licenses": "/dtokils",
            "uninstall_installed_token_activation_issuance_license": "/rtokil:value",
            "set_token_based_activation_flag": "/stokflag",
            "clear_token_based_activation_flag": "/ctokflag",
            "display_token_based_activation_certificates": "/dtokcerts",
            "token_activate": "/tokact:value1:value2",
        }

    def find_ospp(self):
        # Define the list of Office versions (4, 5, and 6 in this case)
        office_versions = [4, 5, 6]

        for version in office_versions:
            # Check if the ospp.vbs file exists in the 32-bit Program Files directory
            program_files_path = os.environ.get("ProgramFiles")
            ospp_vbs_path = os.path.join(program_files_path, f"Microsoft Office\\Office1{version}\\ospp.vbs")

            if os.path.exists(ospp_vbs_path):
                return ospp_vbs_path

            # Check if the ospp.vbs file exists in the 64-bit Program Files directory
            program_files_x86_path = os.environ.get("ProgramFiles(x86)")
            ospp_vbs_path_x86 = os.path.join(program_files_x86_path, f"Microsoft Office\\Office1{version}\\ospp.vbs")

            if os.path.exists(ospp_vbs_path_x86):
                return ospp_vbs_path_x86

        return None

    def execute_ospp_command(self, command):
        ospp_path = self.find_ospp()
        if not ospp_path:
            raise FileNotFoundError("--- ERROR --- \n -Errmsg: ospp.vbs not found. Please check the path.")

        try:
            ospp_command = ["cscript", ospp_path, command]
            result = subprocess.run(ospp_command, capture_output=True, text=True)
            return result.stdout.strip()
        except subprocess.CalledProcessError as e:
            print("--- ERROR --- \n -Errmsg: Error executing ospp.vbs:", e)

    def run_ospp_command_user_input(self, command):
       
        if command in self.ospp_commands:
            ospp_command = self.ospp_commands[command]
            if ":value" in ospp_command:
                value = input(f"Enter the value for {ospp_command.split(':')[0]}: ").strip()
                if not value:
                    raise ValueError("--- ERROR --- \n -Errmsg: Value cannot be empty. Please try again.")
                if 'inpkey' in ospp_command:
                    #remove any dashes
                    value = value.replace('-','').replace(' ', '')
                    value = '-'.join(value[i:i+5] for i in range(0, len(value), 5))
                ospp_command = ospp_command.replace(":value", f":{value}")
            return self.execute_ospp_command(ospp_command)
        else:
            raise ValueError("--- ERROR --- \n -Errmsg: Invalid command.")
        
    def run_ospp_command(self, command, value=None):
        if command in self.ospp_commands:
            ospp_command = self.ospp_commands[command]
            if ":value" in ospp_command:
                if not value:
                    raise ValueError("--- ERROR --- \n -Errmsg: Command needs a value to proceed. Please provide a valid value")
                ospp_command = ospp_command.replace(":value", f":{value}")
            return self.execute_ospp_command(ospp_command)
        else:
            raise ValueError("--- ERROR --- \n -Errmsg: Invalid command.")
    

class OfficeActivation:
    def __init__(self) -> None:
        self.config = ConfigParser()
        self.config.read('config.ini')
        self.ospp = OSPPModule()
        self.actions = {
            'activate_office_with_product_key': self.pid_activation,
            'activate_office_with_installation_id': self.iid_activation,
            'activate_office_with_confirmation_id': self.cid_activation,
        }
        self.ospp_actions = {
            'display_installation_id': self.ospp.run_ospp_command,
            'display_license_info': self.ospp.run_ospp_command
        }
        self.iid = ''
        self.cid = ''
        self.api_key = self.config.get('user_config', 'api_key')
    
    @property
    def api_url(self):
        return f'http://getcid.info/api/{self.iid}/{self.api_key}'
    
    def run_actions(self, action):
        if action in self.actions:
            return self.actions[action]()
        elif action in self.ospp_actions:
            print(self.ospp_actions[action](action))
    
        
    def print_available_options(self):
        print("Available options for Office Activation:")
        for idx, option in enumerate(self.actions.keys(), 1):
            print(f"{idx}. {option}")
        
        print('\nOptions for displaying Office info:')
        for idx, option in enumerate(self.ospp_actions.keys(), 4):
            print(f"{idx}. {option}")


    def get_action_by_number(self, number):
        if number < 1 or number > len(self.actions)+len(self.ospp_actions):
            raise ValueError("--- ERROR --- \n -Errmsg: Invalid input. Please enter a valid number or 'exit' to quit \n-------------")
        
        return list(self.actions.keys() if number < 4 else self.ospp_actions)[number - 1 if number < 4 else number - 4]
        
    def get_installation_id(self):
        try:
            output = self.ospp.run_ospp_command('display_installation_id')
            # Use regex to find the lines containing "Installation ID for"
            pattern = r"Installation ID for:[^\n]*"
            matches = re.findall(pattern, output)

            installation_ids = []
            for match in matches:
                # Check if the line contains "Retail edition" or "MSDNR_Retail edition"
                if "Retail edition" in match or "MSDNR_Retail edition" in match:
                    # Extract the installation ID using regex
                    id_pattern = r"(?<=: )(\d+)(?=[^:]*$)"
                    id_match = re.search(id_pattern, match)
                    if id_match:
                        installation_id = id_match.group(1)
                        installation_ids.append(installation_id)
        
            if installation_ids and len(installation_ids) == 1:
                return installation_ids[0]
            else:
                raise Exception("--- ERROR --- \n -Errmsg: Installation ID not found/multiple IDs found in the output \n -------------")

        except subprocess.CalledProcessError as e:
            print("--- ERROR --- \n -Errmsg: Error retrieving installation ID:", e, ' \n -------------')
    
    def pid_activation(self):
        #install product key
        install_prod_key_command = self.ospp.run_ospp_command_user_input('install_product_key')
        print(install_prod_key_command)
        self.iid = self.get_installation_id()
        if not self.iid:
            raise ValueError('--- ERROR --- \n -Errmsg: No Installation ID returned \n -------------')
        
        #make request for confirmation id
        response = requests.get(self.api_url)
        response.raise_for_status()
        self.cid = response.json()
        with open('cid.txt', 'w') as file:
            # Write cid to file incase of failure
            print('------------------------------------ Outputting Confirmation ID to cid.txt --------------------------------------')
            print('--- In the event the cid activation failed, please proceed to option 3 and key in the cid from the text field ---')
            print('-----------------------------------------------------------------------------------------------------------------')
            file.write(str(self.cid))
        
        if not self.cid:
            raise ValueError('--- ERROR --- \n -Errmsg: Confirmation ID cannot be empty \n -------------')
        activate = self.ospp.run_ospp_command('activate_with_confirmation_id', self.cid)
        print(activate)
                
    def iid_activation(self):
        self.iid = input('Please enter the Installation ID: ')
        if not self.iid or len(self.iid) != 63:
            raise ValueError('--- ERROR --- \n -Errmsg: Please enter a valid Installation ID \n -------------')
        
        #make request for confirmation id
        response = requests.get(self.api_url)
        response.raise_for_status()
        self.cid = response.json()
        with open('cid.txt', 'w') as file:
            # Write cid to file incase of failure
            file.write(str(self.cid))
        
        if not self.cid:
            raise ValueError('--- ERROR --- \n -Errmsg: Confirmation ID cannot be empty \n -------------')
        activate = self.ospp.run_ospp_command('activate_with_confirmation_id', self.cid)
        print(activate)

    def cid_activation(self):
        self.cid = input('Please enter the Confirmation ID: ')
        
        if not self.cid or len(self.cid) != 48:
            raise ValueError('--- ERROR --- \n -Errmsg: Please enter a valid Confirmation ID \n -------------')
        activate = self.ospp.run_ospp_command('activate_with_confirmation_id', self.cid)
        print(activate)
        
    def start(self):
        print("--- Welcome to the Office Activation Script! ---")
        while True:
            self.print_available_options()
            user_input = input("Enter the number of the command you want to run (or 'exit' to quit): ").lower()

            if user_input == "exit":
                break

            try:
                action_number = int(user_input)
                action = self.get_action_by_number(action_number)
                if action:
                    self.run_actions(action)
                else:
                    print("--- ERROR --- \n -Errmsg: Invalid command number. \n -------------")
            except ValueError as e:
                print(str(e))
            except Exception as e:
                traceback.print_exc()
                print(str(e))


if __name__ == "__main__":
    activation_script = OfficeActivation()
    activation_script.start()
