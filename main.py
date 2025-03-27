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
import importlib.util
from typing import Optional, Tuple

CURRENT_VERSION = 'v.1.1.8'
REPO_URL = 'api.github.com'
RELEASE_PATH = f'/repos/lukash333/SAPp-Opener/releases/latest'

active_instances = {}

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
                'tool_width': '300',
            },
            'DEFAULT_SAP_CLIENT': {
                'QG1': '200'
            },
            'APP': {
                'excel': r'C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE'
            },
            'WEB': {
                'w': 'https://pl.wikipedia.org/wiki/'
            },
            'PY': {}
        }

    def _load_config(self) -> None:
        """Load the config file."""
        self.config.read(self.config_file)

    def write_position(self, x: int, y: int, key_x: str, key_y: str) -> None:
        """Save window position to config."""
        self.config['DEFAULT'][key_x] = str(x)
        self.config['DEFAULT'][key_y] = str(y)
        self.save()

    def save(self) -> None:
        """Save the configuration file."""
        with self.config_file.open('w') as configfile:
            self.config.write(configfile)

    def get_position(self, key_x: str, key_y: str) -> Tuple[int, int]:
        """Retrieve the saved window position."""
        x = self.config['DEFAULT'].getint(key_x, self.config['DEFAULT'].getint('position_x', 0))
        y = self.config['DEFAULT'].getint(key_y, self.config['DEFAULT'].getint('position_y', 0))
        
        screen_width, screen_height = self.get_screen_size()

        if x > screen_width:
            x = screen_width - 200

        if y > screen_height:
            y = screen_height - 200

        return x, y

    def get_screen_size(self) -> Tuple[int, int]:
        """Get screen size using Windows API."""
        user32 = ctypes.windll.user32
        return user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)

    def get_def_client(self, system: str) -> Optional[str]:
        """Return the client for the given system from the DEFAULT_SAP_CLIENT section."""
        return self.config['DEFAULT_SAP_CLIENT'].get(system, None)
    
    def get_tool_width(self) -> int:
        return self.config.get('DEFAULT','tool_width')

    def find_sapshcut_exe(self) -> Optional[str]:
        """Search for sapshcut.exe in common directories."""
        if self.config.has_option('DEFAULT','sapshcut_path'):
            return self.config.get('DEFAULT','sapshcut_path')
        
        search_dirs = [pathlib.Path(r"C:\Program Files"), pathlib.Path(r"C:\Program Files (x86)")]

        for directory in search_dirs:
            for path in directory.rglob('sapshcut.exe'):
                return str(path)
        return "None"
    
    def get_path(self, shortcut: str) -> Optional[Tuple[str,str]]:
        """Retrieve path and section type for the given shortcut."""
        for section in ['APP', 'WEB', 'PY']:
            if self.config.has_option(section, shortcut):
                return self.config.get(section, shortcut), section
        return None

class Window:
    def __init__(self, root: tk.Tk, config_manager: ConfigManager):
        self.config_manager = config_manager
        self.move_var = tk.BooleanVar(value=False)

        self.prev_width = root.winfo_screenwidth()
        self.prev_height = root.winfo_screenheight()

        self.root = root
        self.setup_window()
        self.load_window_position()
        self.create_widgets()
        self.bind_events()
        self.start_bring_to_front()
        self.check_update()
        self.check_resolution_change()

    def setup_window(self) -> None:
        """Configure the main window properties."""
        self.root.title("Simple Widget App")
        self.root.geometry(f"{self.config_manager.get_tool_width()}x100")
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
        self.context_menu.add_command(label="Config", command=self.config_open)

    def run_update(self) -> None:
        updater.update_application()

    def config_open(self) -> None:
        subprocess.run(["notepad", "config.ini"])

    def reload(self) -> None:
        subprocess.Popen(['python', 'main.py'], creationflags=subprocess.CREATE_NO_WINDOW)
        os._exit(0)  # Exit the current process

    def check_update(self) -> None:
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
        self.root.after_cancel(self.bring_to_front_id)

        menu_x = self.root.winfo_rootx() + self.root.winfo_width() - 50
        menu_y = self.root.winfo_rooty()

        self.context_menu.post(menu_x, menu_y)     
        self.root.after(100, self.check_menu_closed)

    def check_menu_closed(self) -> None:
        if not self.context_menu.winfo_ismapped():
            self.start_bring_to_front()  # Restart bring-to-front loop
        else:
            self.root.after(100, self.check_menu_closed)  # Keep checking       

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
    
    def check_resolution_change(self) -> None:
        current_width = self.root.winfo_screenwidth()
        current_height = self.root.winfo_screenheight()

        if (current_width, current_height) != (self.prev_width, self.prev_height):
            self.prev_width, self.prev_height = current_width, current_height
            self.load_window_position()

        self.root.after(1000, self.check_resolution_change)

    def start_bring_to_front(self) -> None:
        """Keep bringing the window to the front."""
        self.root.after(500, self.bring_to_front)

    def bring_to_front(self) -> None:
        """Bring the window to the front."""
        self.root.lift()
        self.bring_to_front_id = self.root.after(500, self.bring_to_front)

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
            self.run_application(link)
        elif link_type  == 'WEB':
            self.open_webpage(link)
        elif link_type == 'PY':
            self.run_python(link)
        else:
            print('App type not recognized')

    def run_python(self, py_file: str) -> None:

        class_name = py_file[:-3].capitalize()

        if os.path.exists(py_file):
            try:
                spec = importlib.util.spec_from_file_location(class_name, py_file)
                module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(module)
                cls = getattr(module, class_name)
            except Exception as e:
                print(f"Error importing class from {py_file}: {e}")
                return
        else:
            print(f"Python file {py_file} not found!")
        
        if cls:
            if class_name in active_instances:
                instance = active_instances[class_name]
                if hasattr(instance, 'stop'):
                    instance.stop()
                    del active_instances[class_name]
                else:
                    instance.run()
            else:
                instance = cls()
                instance.run()
                active_instances[class_name] = instance
        else: 
            print(f"Failed to load class {class_name} from {py_file}")


    def run_application(self, link: str) -> None:
        try:
            subprocess.Popen(link)
        except FileNotFoundError as e:
            print(f"Error: {e}")

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
        try:
            """Fetches the latest release information from GitHub."""
            conn = http.client.HTTPSConnection(REPO_URL)
            headers = {'User-Agent': 'SAPp-Opener'}
            conn.request("GET", RELEASE_PATH, headers=headers)
            response = conn.getresponse()
            
            if response.status != 200:
                raise Exception(f"Failed to fetch release info: {response.status}")
            
            data = response.read()
            conn.close()
            return json.loads(data)
        except (http.client.HTTPException, ConnectionError, OSError) as e:
            print(f"Warning: Could not check for updates due to network error: {e}")
            return None 
    
    def download_file(self, file_url: str, local_filename: str) -> None:
        try:
            with urllib.request.urlopen(file_url) as response:
                with open(local_filename, 'wb') as f:
                    f.write(response.read())
        except Exception as e:    
            print(f'An error occurred: {e}')
            
    def update_application(self):
        
        has_update, latest_release, latest_version = self.check_update()

        if has_update:
            print(f"Updating from version {CURRENT_VERSION} to {latest_version}...")
            for asset in latest_release['assets']:
                if asset['name'].endswith('.py'):
                    self.download_file(asset['browser_download_url'], asset['name'])
                    print(f"Downloaded {asset['name']}")

            print("Update completed. Restarting the application...")
            
            subprocess.Popen(['python', 'main.py'], creationflags=subprocess.CREATE_NO_WINDOW)
            os._exit(0) 
        else:
            print(f"Current and Update versions are the same {latest_version}")

    def check_update(self):

        import re

        latest_release = self.get_latest_release_info()
        if latest_release is None:
            print("Skipping update check due to network issues.")
            return False, None, None
        latest_version = latest_release['tag_name']

        latest_version_tuple = tuple(map(int, re.findall(r'\d+', latest_version)))
        current_version_tuple = tuple(map(int, re.findall(r'\d+', CURRENT_VERSION)))
        
        if latest_version_tuple > current_version_tuple:
            return True, latest_release, latest_version
        return False, None, None

if __name__ == "__main__":
    config_manager = ConfigManager()
    updater = Updater()

    root = tk.Tk()
    app = Window(root, config_manager)
    root.mainloop()
