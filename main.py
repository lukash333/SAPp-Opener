import configparser
import pathlib
import subprocess
import sys
import ctypes
import os
import http.client
import json
import tkinter as tk
import urllib.request
from typing import Optional, Tuple

CURRENT_VERSION = 'v.1.1.0'
REPO_URL = 'api.github.com'
RELEASE_PATH = f'/repos/lukash333/SAPp-Opener/releases/latest'

class ConfigManager:
    """Handles reading, writing, and accessing configuration settings."""

    def __init__(self, config_file: str ='config.ini'):
        self.config_file = pathlib.Path(config_file)
        self.config = configparser.ConfigParser()

        if not self.config_file.exists():
            self._create_default_config()

        self._load_config()
        self._merge_default_config()
        self._load_config()

        self.sappath = self.config['DEFAULT'].get('sapshcut_path') or self.find_sapshcut_exe()
        self.default_lang = self.config['DEFAULT'].get('default_sap_lang', 'EN')

    def _create_default_config(self) -> None:
        """Creates a default configuration file if it doesn't exist."""
        default_config = self._default_config()

        for section, options in default_config.items():
            self.config[section] = options

        self.save()
        print(f"Created {self.config_file} with default settings.")

    def _merge_default_config(self) -> None:
        """Merges the default configuration into the existing config."""
        default_config = self._default_config()
        for section, options in default_config.items():
            if section not in self.config:
                self.config[section] = options
            else:
                for key, value in options.items():
                    if key == 'version':
                        self.config[section][key] = value
                    else:
                        if key not in self.config[section]:
                            self.config[section][key] = value
        self.save()

    def _default_config(self) -> dict:
        """Return the default configuration as a dictionary."""
        return {
            'DEFAULT': {
                'app_name': 'SAP Opener',
                'version': CURRENT_VERSION,  # Updated version
                'position_x': '0',
                'position_y': '0',
                'sapshcut_path': self.find_sapshcut_exe(),
                'default_sap_lang': 'EN',
            },
            'DEFAULT_SAP_CLIENT': {
                'QG1': '200'
            },
            'APP': {
                'excel': r'C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE'
            },
            'WEB': {
                'w': 'https://pl.wikipedia.org/wiki/'
            }
        }

    def _load_config(self) -> None:
        """Load the config file."""
        self.config.read(self.config_file)
        print(f"Loaded configuration from {self.config_file}.")

    def write_position(self, x, y, key_x, key_y) -> None:
        """Save window position to config."""
        self.config['DEFAULT'][key_x] = str(x)
        self.config['DEFAULT'][key_y] = str(y)
        self.save()

    def save(self) -> None:
        """Save the configuration file."""
        with self.config_file.open('w') as configfile:
            self.config.write(configfile)

    def get_position(self, key_x, key_y) -> Tuple[int, int]:
        """Retrieve the saved window position."""
        x = self.config['DEFAULT'].getint(key_x, 0)
        y = self.config['DEFAULT'].getint(key_y, 0)
        
        screen_width, screen_height = self.get_screen_size()

        if x > screen_width:
            x = screen_width - 200

        if y > screen_height:
            y = screen_height - 200

        print(screen_width, screen_height)

        return x, y

    def get_screen_size(self) -> Tuple[int, int]:
        """Get screen size using Windows API."""
        user32 = ctypes.windll.user32
        screen_width = user32.GetSystemMetrics(0)  
        screen_height = user32.GetSystemMetrics(1)  
        return screen_width, screen_height

    def get_def_client(self, system: str) -> Optional[str]:
        """Return the client for the given system from the DEFAULT_SAP_CLIENT section."""
        return self.config['DEFAULT_SAP_CLIENT'].get(system, None)

    def find_sapshcut_exe(self) -> Optional[str]:
        """Search for sapshcut.exe in common directories."""
        if self.config.has_option('DEFAULT','sapshcut_path'):
            return self.config.get('DEFAULT','sapshcut_path')
        search_dirs = [pathlib.Path(r"C:\Program Files"), pathlib.Path(r"C:\Program Files (x86)")]

        for directory in search_dirs:
            for path in directory.rglob('sapshcut.exe'):
                print(f"Found sapshcut.exe at: {path}")
                return str(path)
        print("sapshcut.exe not found.")
        return "None"
    
    def get_path(self, shortcut: str) -> Optional[Tuple[str,str]]:
        """Retrieve path and section type for the given shortcut."""
        if self.config.has_option('APP', shortcut):
            return self.config.get('APP', shortcut), 'APP'
        elif self.config.has_option('WEB', shortcut):
            return self.config.get('WEB', shortcut), 'WEB'
        else:
            return None

class Window:
    def __init__(self, root: tk.Tk, config_manager: ConfigManager):
        self.config_manager = config_manager
        self.move_var = tk.BooleanVar(value=False)

        self.root = root
        self.setup_window()
        self.load_window_position()
        self.create_widgets()
        self.bind_events()
        self.start_bring_to_front()
        self.check_update()

    def setup_window(self) -> None:
        """Configure the main window properties."""
        self.root.title("Simple Widget App")
        self.root.geometry("300x100")
        self.root.attributes('-topmost', True)
        self.root.config(bg='magenta')
        self.root.attributes('-transparentcolor', 'magenta')
        self.root.overrideredirect(True)

    def create_widgets(self) -> None:
        """Create and configure UI components."""
        self.entry = tk.Entry(self.root, width=30)
        self.entry.pack(pady=1)

        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_checkbutton(label="Move", variable=self.move_var)
        self.context_menu.add_command(label="Update", command=self.run_update, state="disabled")
        self.context_menu.add_command(label="Exit", command=self.root.destroy)
        self.context_menu.add_command(label="Reload", command=self.reload)

    def run_update(self):

        updater.update_application()

    def reload(self):
        subprocess.Popen(['python', 'main.py'], creationflags=subprocess.CREATE_NO_WINDOW)
        os._exit(0)  # Exit the current process

    def check_update(self):
        has_update, latest_release, latest_version = updater.check_update()

        if has_update:
            self.entry.insert(0, f"Update possible to {latest_version}")
            self.context_menu.entryconfig(1, state="active")

    def bind_events(self) -> None:
        """Bind events to their handlers."""
        self.entry.bind('<Return>', self.on_enter_pressed)
        self.root.bind("<Button-3>", self.show_context_menu)
        self.root.bind('<Button-1>', self.start_move)
        self.root.bind('<B1-Motion>', self.on_motion)

    def on_enter_pressed(self, event) -> None:
        """Handle the Enter key press event."""
        input_text = self.entry.get().strip()
        print(f"You entered: {input_text}")
        InputProcessor(input_text)
        self.entry.delete(0, tk.END)

    def show_context_menu(self, event) -> None:
        """Display the context menu."""
        menu_x = self.root.winfo_x() + self.root.winfo_width() - 50
        menu_y = self.root.winfo_y()
        self.context_menu.post(menu_x, menu_y)            

    def start_move(self, event) -> None:
        """Start window movement."""
        self.root.x, self.root.y = event.x, event.y

    def on_motion(self, event) -> None:
        """Move the window if 'Move' is selected."""
        if self.move_var.get():
            key_x = f"position_x_{root.winfo_screenwidth()}x{root.winfo_screenheight()}"
            key_y = f"position_y_{root.winfo_screenwidth()}x{root.winfo_screenheight()}"
            x = event.x_root - self.root.x
            y = event.y_root - self.root.y
            self.root.geometry(f'+{x}+{y}')
            self.config_manager.write_position(self.root.winfo_x(), self.root.winfo_y(), key_x, key_y)

    def load_window_position(self) -> None:
        """Load and set window position."""

        key_x = f"position_x_{root.winfo_screenwidth()}x{root.winfo_screenheight()}"
        key_y = f"position_y_{root.winfo_screenwidth()}x{root.winfo_screenheight()}"

        x, y = self.config_manager.get_position(key_x, key_y)
        if x or y:
            self.root.geometry(f'+{x}+{y}')

    def start_bring_to_front(self) -> None:
        """Keep bringing the window to the front."""
        self.root.after(500, self.bring_to_front)

    def bring_to_front(self) -> None:
        """Bring the window to the front."""
        self.root.lift()
        self.root.after(500, self.bring_to_front)

class InputProcessor:

    def __init__(self, input_string: str):
        self.input_string = input_string.lower()
        shortcut_details = config_manager.get_path(self.input_string)
        if shortcut_details:
            self.process_configured(shortcut_details)
        else:
            self.process_unconfigured()

    def process_unconfigured(self) -> None:
        """Process SAP input depending on its length."""
        handlers = {
            3: self.run_defaulted,
            5: self.run_with_language,
            6: self.run_with_system_client,
            8: self.run_with_language_client,
        }
        handler = handlers.get(len(self.input_string))
        if handler:
            handler()
    
    def process_configured(self, details: Tuple[str,str]) -> None:
        """Process pre-configured shortcuts."""
        link, link_type  = details
        if link_type  == 'APP':
            try:
                subprocess.Popen(link)
            except FileNotFoundError as e:
                print(f"Error: {e}")
        elif link_type  == 'WEB':
            self.open_webpage(link)
        else:
            print('App type not recognized')

    def open_webpage(self, url: str) -> None:
        """Open a webpage based on the input string."""
        try:
            if sys.platform.startswith('win'):
                subprocess.Popen(['start', url], shell=True)
            elif sys.platform.startswith('darwin'):
                subprocess.Popen(['open', url])
            else:
                subprocess.Popen(['xdg-open', url])
        except Exception as e:
                print(f"An error occurred: {e}")

    def run_defaulted(self) -> None:
        self.run_sap_gui(config_manager.get_def_client(self.input_string), config_manager.default_lang, self.input_string)

    def run_with_language(self) -> None:
        language = self.input_string[:2]
        system = self.input_string[2:]
        self.run_sap_gui(config_manager.get_def_client(system), language, system)

    def run_with_system_client(self) -> None:
        system = self.input_string[:3]
        client = self.input_string[3:]
        self.run_sap_gui(client, config_manager.default_lang, system)

    def run_with_language_client(self) -> None:
        language = self.input_string[:2]
        system = self.input_string[2:5]
        client = self.input_string[5:]
        self.run_sap_gui(client, language, system)

    def run_sap_gui(self, client: Optional[str] = None, language: Optional[str] = None, system: Optional[str] = None, transaction: Optional[str] = None) -> None:
        command = [config_manager.sappath]

        if command == ['None']:
            print(f"SAP Shortcut not installed")
            return
        
        if client:
            command.append(f"-client={client}")
        if language:
            command.append(f"-language={language}")
        if system:
            command.append(f"-system={system}")
        if transaction:
            command.append(f"-transaction={transaction}")

        try:
            subprocess.run(command, check=True)
            print("SAP GUI launched successfully.")
        except subprocess.CalledProcessError as e:
            print(f"Error launching SAP GUI: {e}")

class Updater:
    
    def get_latest_release_info(self):
        """Fetches the latest release information from GitHub."""
        conn = http.client.HTTPSConnection(REPO_URL)
        headers = {
            'User-Agent': 'SAPp-Opener'  # Use your application's name here
        }

        conn.request("GET", RELEASE_PATH, headers=headers)
        response = conn.getresponse()
        
        if response.status != 200:
            raise Exception(f"Failed to fetch release info: {response.status}")
        
        data = response.read()
        conn.close()
        return json.loads(data)
    
    def download_file(self, file_url, local_filename):
        try:
            with urllib.request.urlopen(file_url) as response:
                with open(local_filename, 'wb') as f:
                    f.write(response.read())
            print(f'File downloaded successfully as {local_filename}')
        except Exception as e:    
            print(f'An error occurred: {e}')
            
    def update_application(self):
        
        has_update, latest_release, latest_version = self.check_update()

        if has_update:
            print(f"Updating from version {CURRENT_VERSION} to {latest_version}...")

            # Iterate through the assets and download the .py files
            for asset in latest_release['assets']:
                if asset['name'].endswith('.py'):
                    self.download_file(asset['browser_download_url'], asset['name'])
                    print(f"Downloaded {asset['name']}")

            print("Update completed. Restarting the application...")
            
            # Restart the application
            subprocess.Popen(['python', 'main.py'], creationflags=subprocess.CREATE_NO_WINDOW)
            os._exit(0)  # Exit the current process
        else:
            print(f"Current and Update versions are the same {latest_version}")

    def check_update(self):

        from packaging.version import Version

        latest_release = self.get_latest_release_info()
        latest_version = latest_release['tag_name']
        
        if Version(latest_version[2:]) > Version(CURRENT_VERSION[2:]):
            return True, latest_release, latest_version
        return False, None, None

if __name__ == "__main__":
    config_manager = ConfigManager()
    updater = Updater()

    root = tk.Tk()
    app = Window(root, config_manager)
    root.mainloop()