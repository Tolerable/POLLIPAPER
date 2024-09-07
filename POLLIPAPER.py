import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, colorchooser
import requests
from requests.exceptions import RequestException
import textwrap
from PIL import Image, ImageDraw, ImageFont, UnidentifiedImageError
import io
import re
import random
import threading
from collections import deque
import json
import os
import time
import ctypes
from ctypes import wintypes
import comtypes
from comtypes import GUID
import sys
import pythoncom
import winreg
import piexif

# Define HRESULT
HRESULT = ctypes.HRESULT

# IDesktopWallpaper interface GUID
CLSID_DesktopWallpaper = GUID('{C2CF3110-460E-4fc1-B9D0-8A1C0C9CC4BD}')
IID_IDesktopWallpaper = GUID('{B92B56A9-8B55-4E14-9A89-0199BBB6F93B}')

# Wallpaper position enum
class DESKTOP_WALLPAPER_POSITION(ctypes.c_int):
    DWPOS_CENTER = 0
    DWPOS_TILE = 1
    DWPOS_STRETCH = 2
    DWPOS_FIT = 6
    DWPOS_FILL = 10
    DWPOS_SPAN = 5  # Changed from 22 to 5

class IDesktopWallpaper(comtypes.IUnknown):
    _iid_ = IID_IDesktopWallpaper
    _methods_ = [
        comtypes.COMMETHOD([], HRESULT, 'SetWallpaper',
            (['in'], wintypes.LPCWSTR, 'monitorID'),
            (['in'], wintypes.LPCWSTR, 'wallpaper')),
        comtypes.COMMETHOD([], HRESULT, 'GetWallpaper'),
        comtypes.COMMETHOD([], HRESULT, 'GetMonitorDevicePathAt'),
        comtypes.COMMETHOD([], HRESULT, 'GetMonitorDevicePathCount'),
        comtypes.COMMETHOD([], HRESULT, 'GetMonitorRECT'),
        comtypes.COMMETHOD([], HRESULT, 'SetBackgroundColor'),
        comtypes.COMMETHOD([], HRESULT, 'GetBackgroundColor'),
        comtypes.COMMETHOD([], HRESULT, 'SetPosition',
            (['in'], DESKTOP_WALLPAPER_POSITION, 'position')),
        comtypes.COMMETHOD([], HRESULT, 'GetPosition',
            (['out'], ctypes.POINTER(DESKTOP_WALLPAPER_POSITION), 'position')),
        comtypes.COMMETHOD([], HRESULT, 'SetSlideshow'),
        comtypes.COMMETHOD([], HRESULT, 'GetSlideshow'),
        comtypes.COMMETHOD([], HRESULT, 'SetStatus'),
        comtypes.COMMETHOD([], HRESULT, 'Enable')
    ]

class RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long),
                ("top", ctypes.c_long),
                ("right", ctypes.c_long),
                ("bottom", ctypes.c_long)]

class WeatherFetcher:
    def __init__(self):
        self.api_key = os.environ.get('OPENWEATHER_API_KEY')

    def fetch_weather_data(self, city="Rochester", country="US", units="metric", lang="en"):
        if not self.api_key:
            print("OpenWeather API key not found. Weather features will be disabled.")
            return None

        # Step 1: Get coordinates using Geocoding API
        geocode_url = f"http://api.openweathermap.org/geo/1.0/direct?q={city},{country}&limit=1&appid={self.api_key}"
        geocode_response = requests.get(geocode_url)
        geocode_data = geocode_response.json()

        if not geocode_data:
            raise ValueError(f"Could not find coordinates for {city}, {country}")

        lat = geocode_data[0]['lat']
        lon = geocode_data[0]['lon']

        print(f"Coordinates for {city}, {country}: Lat {lat}, Lon {lon}")

        # Step 2: Fetch weather data using the updated API endpoint
        weather_url = f"https://api.openweathermap.org/data/2.5/weather?lat={lat}&lon={lon}&appid={self.api_key}&units={units}&lang={lang}"
        weather_response = requests.get(weather_url)
        weather_data = weather_response.json()

        # Step 3: Save weather data to JSON file
        with open('weather_data.json', 'w') as f:
            json.dump(weather_data, f, indent=4)

        print("Weather data fetched and saved to weather_data.json")

        return weather_data

# Initialize WeatherFetcher
weather_fetcher = WeatherFetcher()

class PollinationsBackgroundSetter:
    def __init__(self, master):
        self.master = master
        master.title("POLLIPAPER")
        master.geometry("650x600")

        self.enhance = tk.BooleanVar()
        self.wallpaper_style = tk.StringVar(value="fill")

        self.interval = max(300, 60 * 5)  # Minimum 5 minutes
        self.always_on_top = tk.BooleanVar(value=False)
        
        self.model = tk.StringVar(value="flux")
        
        self.use_weather = tk.BooleanVar(value=False)
        self.overlay_weather = tk.BooleanVar(value=False)
        self.overlay_opacity = tk.IntVar(value=128)  # 0-255 for opacity
        self.overlay_position = tk.StringVar(value="top_right")
     
        self.use_drop_shadow = tk.BooleanVar(value=False)
        self.weather_api_key = tk.StringVar()
        self.weather_fetcher = weather_fetcher
        
        self.temp_unit = tk.StringVar(value="F")

        self.overlay_color = tk.StringVar(value="#000000")

        self.submitted_prompt = tk.StringVar()
        self.returned_prompt = tk.StringVar()
        self.negative_prompt = tk.StringVar()
        self.nsfw_status = tk.StringVar()
        self.weather_conditions = tk.StringVar()
        self.current_weather_conditions = "N/A"

        self.is_running = False  # Initialize is_running before calling setup_ui and load_settings
        self.setter_thread = None
        self.current_request_id = None

        self.setup_ui()
        self.load_settings()
        self.update_position_options()
        self.prompt_history = deque(maxlen=20)
        self.load_history()

        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.image_dir = "POLLINATIONS_BACKGROUNDS"
        os.makedirs(self.image_dir, exist_ok=True)

        self.style_dimensions = {
            "center": (1920, 1080),
            "tile": (1920, 1080),
            "stretch": (1920, 1080),
            "fit": (1920, 1080),
            "fill": (1920, 1080),
            "span": (3840, 1080),
        }
        self.wallpaper_style.trace_add('write', self.update_position_options)
        
    def setup_ui(self):
        self.frame = ttk.Frame(self.master, padding="10")
        self.frame.pack(fill=tk.BOTH, expand=True)

        # Prompt input
        prompt_label = ttk.Label(self.frame, text="Prompt:")
        prompt_label.pack(fill=tk.X)

        prompt_frame = ttk.Frame(self.frame)
        prompt_frame.pack(fill=tk.BOTH, expand=True)

        self.prompt_entry = tk.Text(prompt_frame, wrap=tk.WORD, height=4)
        self.prompt_entry.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        prompt_scrollbar = ttk.Scrollbar(prompt_frame, orient="vertical", command=self.prompt_entry.yview)
        prompt_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.prompt_entry.config(yscrollcommand=prompt_scrollbar.set)

        # History dropdown
        history_label = ttk.Label(self.frame, text="History:")
        history_label.pack(fill=tk.X)
        self.history_var = tk.StringVar()
        self.history_dropdown = ttk.Combobox(self.frame, textvariable=self.history_var)
        self.history_dropdown.pack(fill=tk.X)
        self.history_dropdown.bind('<<ComboboxSelected>>', self.on_history_select)

        # Options frame
        options_frame = ttk.LabelFrame(self.frame, text="Options", padding="5")
        options_frame.pack(fill=tk.X, pady=10)

        # First row of options
        first_row = ttk.Frame(options_frame)
        first_row.pack(fill=tk.X)

        self.start_stop_button = ttk.Button(first_row, text="Start", command=self.toggle_start_stop, width=10)
        self.start_stop_button.pack(side=tk.LEFT, padx=(0, 5))

        ttk.Checkbutton(first_row, text="Enhance", variable=self.enhance).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(first_row, text="Wallpaper Style:").pack(side=tk.LEFT)
        style_options = ["fill", "fit", "stretch", "tile", "center", "span"]
        ttk.Combobox(first_row, textvariable=self.wallpaper_style, values=style_options, width=10).pack(side=tk.LEFT)

        ttk.Label(first_row, text="Model:").pack(side=tk.LEFT, padx=(10, 0))
        model_options = ["flux", "flux-realism", "flux-anime", "flux-3d", "any-dark", "turbo"]
        ttk.Combobox(first_row, textvariable=self.model, values=model_options, width=15).pack(side=tk.LEFT)

        # Overlay settings
        overlay_frame = ttk.LabelFrame(self.frame, text="Overlay Settings", padding="5")
        overlay_frame.pack(fill=tk.X, pady=10)

        ttk.Label(overlay_frame, text="Opacity:").pack(side=tk.LEFT)
        ttk.Scale(overlay_frame, from_=0, to=255, variable=self.overlay_opacity, orient=tk.HORIZONTAL, length=100).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(overlay_frame, text="Position:").pack(side=tk.LEFT)
        self.position_dropdown = ttk.Combobox(overlay_frame, textvariable=self.overlay_position, width=15)
        self.position_dropdown.pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(overlay_frame, text="Color:").pack(side=tk.LEFT)
        self.color_button = tk.Button(overlay_frame, width=2, command=self.choose_overlay_color)
        self.color_button.pack(side=tk.LEFT, padx=(0, 10))

        ttk.Checkbutton(overlay_frame, text="Drop Shadow", variable=self.use_drop_shadow).pack(side=tk.LEFT)

        # Add a button in the UI to remove the selected prompt from history
        remove_button = ttk.Button(self.frame, text="Remove from History", command=self.remove_selected_from_history)
        remove_button.pack()

        # Current Image Information frame
        info_frame = ttk.LabelFrame(self.frame, text="Current Image Information", padding="5")
        info_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # Submitted Prompt
        ttk.Label(info_frame, text="Submitted Prompt:").pack(anchor=tk.W)
        self.submitted_prompt_display = tk.Text(info_frame, wrap=tk.WORD, height=2, state='disabled')
        self.submitted_prompt_display.pack(fill=tk.X)

        # Returned Prompt
        ttk.Label(info_frame, text="Returned Prompt:").pack(anchor=tk.W)
        self.returned_prompt_display = tk.Text(info_frame, wrap=tk.WORD, height=2, state='disabled')
        self.returned_prompt_display.pack(fill=tk.X)

        # Weather Conditions
        ttk.Label(info_frame, text="Weather Conditions:").pack(anchor=tk.W)
        self.weather_conditions_display = tk.Text(info_frame, wrap=tk.WORD, height=1, state='disabled')
        self.weather_conditions_display.pack(fill=tk.X)

        # Menu setup
        menubar = tk.Menu(self.master)
        options_menu = tk.Menu(menubar, tearoff=0)
        options_menu.add_checkbutton(label="Always on Top", variable=self.always_on_top, command=self.toggle_always_on_top)
        options_menu.add_command(label="Set Interval (min)", command=self.set_interval)
        options_menu.add_checkbutton(label="Use Weather-Based Prompts", variable=self.use_weather)
        options_menu.add_checkbutton(label="Overlay Weather Info", variable=self.overlay_weather)
        options_menu.add_command(label="Set OpenWeatherMap API Key", command=self.set_weather_api_key)

        # Temperature unit submenu
        temp_unit_menu = tk.Menu(options_menu, tearoff=0)
        temp_unit_menu.add_radiobutton(label="Fahrenheit (°F)", variable=self.temp_unit, value="F")
        temp_unit_menu.add_radiobutton(label="Celsius (°C)", variable=self.temp_unit, value="C")
        options_menu.add_cascade(label="Temperature Unit", menu=temp_unit_menu)

        menubar.add_cascade(label="Options", menu=options_menu)
        self.master.config(menu=menubar)

        self.update_position_options()
        self.update_color_button()
    
        self.current_prompt_label = ttk.Label(self.frame, text="Current Prompt:")
        self.current_prompt_label.pack(fill=tk.X, pady=(10, 0))
        
        self.current_prompt_display = tk.Text(self.frame, wrap=tk.WORD, height=3, state='disabled')
        self.current_prompt_display.pack(fill=tk.X)

    def cleanup_background_images(self):
        backgrounds = sorted(
            [f for f in os.listdir(self.image_dir) if f.startswith("background_") and f.endswith(".png")],
            key=lambda x: os.path.getmtime(os.path.join(self.image_dir, x)),
            reverse=True
        )
        
        wallpapers = sorted(
            [f for f in os.listdir(self.image_dir) if f.startswith("wallpaper_") and f.endswith(".png")],
            key=lambda x: os.path.getmtime(os.path.join(self.image_dir, x)),
            reverse=True
        )
        
        # Keep only the 5 most recent background images
        for old_bg in backgrounds[5:]:
            try:
                os.remove(os.path.join(self.image_dir, old_bg))
                print(f"Deleted old background: {old_bg}")
            except Exception as e:
                print(f"Error deleting {old_bg}: {e}")
        
        # Keep only the most recent wallpaper image
        for old_wp in wallpapers[1:]:
            try:
                os.remove(os.path.join(self.image_dir, old_wp))
                print(f"Deleted old wallpaper: {old_wp}")
            except Exception as e:
                print(f"Error deleting {old_wp}: {e}")
        
        print(f"After cleanup: {len(backgrounds[:5])} backgrounds and {len(wallpapers[:1])} wallpapers kept.")

    def update_position_options(self, *args):
        if self.wallpaper_style.get() == "span":
            options = ["left_top_left", "left_top_right", "left_bottom_left", "left_bottom_right",
                       "right_top_left", "right_top_right", "right_bottom_left", "right_bottom_right"]
        else:
            options = ["top_left", "top_right", "bottom_left", "bottom_right"]
        
        self.position_dropdown['values'] = options
        if self.overlay_position.get() not in options:
            self.overlay_position.set(options[0])  # Set to first option if current is not valid

    def choose_overlay_color(self):
        color = colorchooser.askcolor(title="Choose overlay background color")
        if color[1]:  # color is None if dialog is cancelled
            self.overlay_color.set(color[1])
            self.update_color_button()

    def update_color_button(self):
        if hasattr(self, 'color_button'):
            self.color_button.config(bg=self.overlay_color.get())

    def update_current_prompt_display(self, prompt):
        try:
            if hasattr(self, 'current_prompt_display'):
                self.current_prompt_display.config(state='normal')
                self.current_prompt_display.delete(1.0, tk.END)
                self.current_prompt_display.insert(tk.END, prompt)
                self.current_prompt_display.config(state='disabled')
            else:
                print("Error: current_prompt_display widget not found")
        except Exception as e:
            print(f"Error updating current prompt display: {e}")

    def update_prompt_info(self, original_prompt, returned_prompt):
        print(f"Updating prompts - Original: {original_prompt}")
        print(f"Full returned data: {returned_prompt}")
        
        self.submitted_prompt.set(original_prompt)
        if self.enhance.get():
            # Try to extract the enhanced prompt
            enhanced_prompt = re.search(r'"([^"]*)"', returned_prompt)
            if enhanced_prompt:
                cleaned_prompt = enhanced_prompt.group(1)
            else:
                # If no quotes found, use the entire returned prompt
                cleaned_prompt = returned_prompt
            
            # Remove the original prompt if it appears at the end of the returned prompt
            cleaned_prompt = cleaned_prompt.replace(original_prompt.strip(), '').strip()
            
            print(f"Cleaned enhanced prompt: {cleaned_prompt}")
            self.returned_prompt.set(cleaned_prompt)
        else:
            self.returned_prompt.set(original_prompt)
        
        # Update the Text widgets
        if hasattr(self, 'submitted_prompt_display'):
            self.submitted_prompt_display.config(state='normal')
            self.submitted_prompt_display.delete(1.0, tk.END)
            self.submitted_prompt_display.insert(tk.END, self.submitted_prompt.get())
            self.submitted_prompt_display.config(state='disabled')
        
        if hasattr(self, 'returned_prompt_display'):
            self.returned_prompt_display.config(state='normal')
            self.returned_prompt_display.delete(1.0, tk.END)
            self.returned_prompt_display.insert(tk.END, self.returned_prompt.get())
            self.returned_prompt_display.config(state='disabled')
        
        if self.use_weather.get():
            self.weather_conditions.set(self.current_weather_conditions)
        else:
            self.weather_conditions.set('N/A')
        
        # Update weather conditions display
        if hasattr(self, 'weather_conditions_display'):
            self.weather_conditions_display.config(state='normal')
            self.weather_conditions_display.delete(1.0, tk.END)
            self.weather_conditions_display.insert(tk.END, self.weather_conditions.get())
            self.weather_conditions_display.config(state='disabled')

    def set_weather_api_key(self):
        key = simpledialog.askstring("API Key", "Enter your OpenWeatherMap API Key:", show='*')
        if key:
            self.weather_api_key.set(key)
            # Show only last 4 digits
            masked_key = '*' * (len(key) - 4) + key[-4:]
            print(f"API Key set: {masked_key}")
            # Update environment variable and save settings
            self.update_api_key(key)

    def update_api_key(self, key):
        os.environ['OPENWEATHER_API_KEY'] = key
        self.save_settings()
        self.weather_fetcher = WeatherFetcher()

    def set_interval(self):
        interval_str = simpledialog.askstring("Set Interval", "Enter interval in minutes:", initialvalue=str(self.interval / 60), parent=self.master)
        try:
            self.interval = max(0.1, float(interval_str)) * 60
            print(f"Interval set to {self.interval} seconds.")
            self.save_settings()
        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter a valid number for the interval.")

    def toggle_always_on_top(self):
        self.master.wm_attributes("-topmost", self.always_on_top.get())
        self.save_settings()

    def toggle_start_stop(self):
        if self.is_running:
            self.stop_setter()
        else:
            self.start_setter()

    def on_history_select(self, event):
        selected_prompt = self.history_var.get()
        self.prompt_entry.delete(1.0, tk.END)
        self.prompt_entry.insert(tk.END, selected_prompt)

    def add_to_history(self, prompt):
        if prompt:
            # Remove the prompt if it already exists in the history
            self.prompt_history = deque([p for p in self.prompt_history if p != prompt], maxlen=20)
            # Add the prompt to the beginning of the history
            self.prompt_history.appendleft(prompt)
            self.update_history_dropdown()
            self.save_history()
            print(f"Added to history: {prompt}")  # Debug print

    def remove_from_history(self, prompt):
        if prompt in self.prompt_history:
            self.prompt_history.remove(prompt)
            self.update_history_dropdown()
            self.save_history()

    def remove_selected_from_history(self):
        selected = self.history_var.get()
        if selected:
            self.remove_from_history(selected)

    def update_history_dropdown(self):
        history_list = list(self.prompt_history)
        self.history_dropdown['values'] = history_list
        if history_list:
            self.history_var.set(history_list[0])
        print(f"Updated history dropdown. Current history: {history_list}")  # Debug print

    def save_history(self):
        with open('prompt_history.json', 'w') as f:
            json.dump(list(self.prompt_history), f)
        print(f"Saved history: {list(self.prompt_history)}")  # Debug print

    def load_history(self):
        if os.path.exists('prompt_history.json'):
            with open('prompt_history.json', 'r') as f:
                self.prompt_history = deque(json.load(f), maxlen=20)
        else:
            self.prompt_history = deque(maxlen=20)
        self.update_history_dropdown()
        print(f"Loaded history: {list(self.prompt_history)}")  # Debug print

    def start_setter(self):
        if self.use_weather.get():
            weather_data = self.weather_fetcher.fetch_weather_data()
            if weather_data:
                prompt = self.generate_weather_prompt(weather_data)
                try:
                    self.update_current_prompt_display(prompt)
                except Exception as e:
                    print(f"Error updating current prompt display: {e}")
            else:
                messagebox.showerror("Weather Data Error", "Failed to fetch weather data.")
                return
        else:
            prompt = self.prompt_entry.get("1.0", tk.END).strip()
            if not prompt:
                if self.prompt_history:
                    prompt = self.prompt_history[0]
                    self.prompt_entry.insert(tk.END, prompt)
                else:
                    messagebox.showerror("Missing Prompt", "Please enter a prompt or ensure there is one in history.")
                    return
            self.add_to_history(prompt)  # Only add to history if it's not a weather-based prompt

        try:
            self.update_current_prompt_display(prompt)
        except Exception as e:
            print(f"Error updating current prompt display: {e}")

        self.is_running = True
        self.start_stop_button.config(text="Stop")
        self.current_request_id = random.randint(1, 1000000)
        
        self.setter_thread = threading.Thread(target=self.run_setter, args=(prompt, self.interval, self.current_request_id))
        self.setter_thread.start()

    def stop_setter(self):
        self.is_running = False
        self.current_request_id = None  # This will invalidate any ongoing requests
        if self.setter_thread and self.setter_thread.is_alive():
            self.setter_thread.join(timeout=1)  # Wait for the thread to finish, but not indefinitely
        self.start_stop_button.config(text="Start")

    def run_setter(self, prompt, interval, request_id):
        while self.is_running and request_id == self.current_request_id:
            try:
                self.fetch_and_set_background(prompt, request_id)
            except Exception as e:
                print(f"Error setting background: {e}")
            time.sleep(interval)

    def get_image_dimensions(self, style):
        return self.style_dimensions.get(style, (1920, 1080))

    def fetch_and_set_background(self, prompt, request_id):
        max_retries = 5
        retry_delay = 5
        for attempt in range(max_retries):
            if not self.is_running or request_id != self.current_request_id:
                print("Operation cancelled.")
                return
            try:
                seed = random.randint(1, 1000000)
                enhance_param = "true" if self.enhance.get() else "false"
                style = self.wallpaper_style.get()
                
                width, height = self.get_image_dimensions(style)
                print(f"Style: {style}, Dimensions: {width}x{height}")
                
                weather_data = None
                if self.overlay_weather.get() or self.use_weather.get():
                    weather_data = self.weather_fetcher.fetch_weather_data()
                    if not weather_data:
                        print("Failed to fetch weather data.")
                        if self.use_weather.get():
                            print("Cannot generate weather-based image. Using user prompt.")
                            self.use_weather.set(False)  # Temporarily disable weather-based prompt
                
                if self.use_weather.get() and weather_data:
                    prompt = self.generate_weather_prompt(weather_data)
                    print(f"Using weather-based prompt: {prompt}")
                else:
                    print(f"Using user prompt: {prompt}")
                
                url = f"https://image.pollinations.ai/prompt/{prompt}?nologo=true&nofeed=true&model={self.model.get()}&enhance={enhance_param}&seed={seed}&width={width}&height={height}"
                print(f"Attempt {attempt + 1}: Fetching image with URL: {url}")
                
                response = requests.get(url, timeout=45)
                
                if response.status_code == 200 and request_id == self.current_request_id:
                    print("Image fetched successfully.")
                    image_data = response.content
                    
                    try:
                        image = Image.open(io.BytesIO(image_data))
                        actual_width, actual_height = image.size
                        print(f"Actual image dimensions: {actual_width}x{actual_height}")
                        
                        # Extract metadata from EXIF
                        exif_data = image.info.get('exif', b'')
                        if exif_data and self.enhance.get():
                            # Find the JSON string in the EXIF data
                            json_match = re.search(b'{"prompt":.*}', exif_data)
                            if json_match:
                                json_str = json_match.group(0).decode('utf-8')
                                try:
                                    metadata_dict = json.loads(json_str)
                                    returned_prompt = metadata_dict.get('prompt', 'Unable to retrieve returned prompt')
                                    self.update_prompt_info(prompt, returned_prompt)
                                except json.JSONDecodeError as e:
                                    print(f"Failed to parse JSON in metadata: {e}")
                                    self.update_prompt_info(prompt, prompt)
                            else:
                                print("No JSON data found in EXIF")
                                self.update_prompt_info(prompt, prompt)
                        else:
                            self.update_prompt_info(prompt, prompt)
                        
                        if (actual_width, actual_height) != (width, height):
                            print(f"Warning: Received image dimensions ({actual_width}x{actual_height}) "
                                  f"do not match requested dimensions ({width}x{height})")
                    except UnidentifiedImageError:
                        print(f"Error: Received data is not a valid image. Retrying...")
                        raise  # This will trigger the retry mechanism
                
                    timestamp = int(time.time())
                    image_path = os.path.join(self.image_dir, f"background_{timestamp}.png")
                    image.save(image_path, 'PNG')
                    print(f"Image saved to {image_path}")
                    
                    if self.overlay_weather.get() and weather_data:
                        try:
                            overlay_image = self.apply_weather_overlay(image.copy(), weather_data)
                            wallpaper_path = os.path.join(self.image_dir, f"wallpaper_{timestamp}.png")
                            overlay_image.save(wallpaper_path)
                            print(f"Weather overlay added to {wallpaper_path}")
                            self.set_windows_background(wallpaper_path)
                        except Exception as e:
                            print(f"Error applying weather overlay: {e}")
                            self.set_windows_background(image_path)
                    else:
                        self.set_windows_background(image_path)
                    
                    self.cleanup_background_images()
                    return
                else:
                    print(f"Failed to fetch image. Status code: {response.status_code}")
            
            except (RequestException, UnidentifiedImageError) as e:
                print(f"Error on attempt {attempt + 1}: {e}")
            except Exception as e:
                print(f"Unexpected error on attempt {attempt + 1}: {e}")
            
            if attempt < max_retries - 1:
                print(f"Retrying in {retry_delay} seconds...")
                time.sleep(retry_delay)
                retry_delay *= 2
        
        print("Failed to fetch and set background after all retries.")
    
    def save_settings(self):
        settings = {
            "enhance": self.enhance.get(),
            "wallpaper_style": self.wallpaper_style.get(),
            "interval": self.interval,
            "always_on_top": self.always_on_top.get(),
            "model": self.model.get(),
            "use_weather": self.use_weather.get(),
            "overlay_weather": self.overlay_weather.get(),
            "overlay_opacity": self.overlay_opacity.get(),
            "overlay_position": self.overlay_position.get(),
            "use_drop_shadow": self.use_drop_shadow.get(),
            "weather_api_key": self.weather_api_key.get(),
            "temp_unit": self.temp_unit.get(),
            "overlay_color": self.overlay_color.get()
        }
        with open('background_settings.json', 'w') as f:
            json.dump(settings, f)

    def load_settings(self):
        if os.path.exists('background_settings.json'):
            with open('background_settings.json', 'r') as f:
                settings = json.load(f)
                self.enhance.set(settings.get("enhance", False))
                self.wallpaper_style.set(settings.get("wallpaper_style", "fill"))
                self.interval = max(300, settings.get("interval", 300))
                self.always_on_top.set(settings.get("always_on_top", False))
                self.model.set(settings.get("model", "flux"))
                self.use_weather.set(settings.get("use_weather", False))
                self.overlay_weather.set(settings.get("overlay_weather", False))
                self.overlay_opacity.set(settings.get("overlay_opacity", 128))
                self.overlay_position.set(settings.get("overlay_position", "top_right"))
                self.use_drop_shadow.set(settings.get("use_drop_shadow", False))
                self.weather_api_key.set(settings.get("weather_api_key", ""))
                self.temp_unit.set(settings.get("temp_unit", "F"))
                self.overlay_color.set(settings.get("overlay_color", "#000000"))
                
                if self.weather_api_key.get():
                    os.environ['OPENWEATHER_API_KEY'] = self.weather_api_key.get()
                
                self.update_position_options()
                self.update_color_button()
                self.toggle_always_on_top()

        # Apply loaded settings to the UI
        self.apply_loaded_settings()

    def set_windows_background(self, image_path):
        try:
            pythoncom.CoInitialize()
            abs_path = os.path.abspath(image_path)
            if not os.path.exists(abs_path):
                raise FileNotFoundError(f"The file {abs_path} does not exist.")

            style = self.wallpaper_style.get()
            
            idw = comtypes.CoCreateInstance(CLSID_DesktopWallpaper, interface=IDesktopWallpaper)
            
            # Set the wallpaper
            try:
                idw.SetWallpaper(None, abs_path)
                print(f"Wallpaper set to {abs_path}")
            except Exception as e:
                print(f"Error setting wallpaper: {e}")
            
            # Set the style
            try:
                style_map = {
                    "center": DESKTOP_WALLPAPER_POSITION.DWPOS_CENTER,
                    "tile": DESKTOP_WALLPAPER_POSITION.DWPOS_TILE,
                    "stretch": DESKTOP_WALLPAPER_POSITION.DWPOS_STRETCH,
                    "fit": DESKTOP_WALLPAPER_POSITION.DWPOS_FIT,
                    "fill": DESKTOP_WALLPAPER_POSITION.DWPOS_FILL,
                    "span": DESKTOP_WALLPAPER_POSITION.DWPOS_SPAN
                }
                idw.SetPosition(style_map.get(style, DESKTOP_WALLPAPER_POSITION.DWPOS_FILL))
                print(f"Wallpaper style set to {style}")
            except Exception as e:
                print(f"Error setting wallpaper style: {e}")
            
        except Exception as e:
            print(f"Unexpected error in set_windows_background: {e}")
        finally:
            pythoncom.CoUninitialize()

    def generate_weather_prompt(self, weather_data):
        weather_condition = weather_data['weather'][0]['main'].lower()
        current_time = time.localtime()
        
        # Determine time of day
        if 5 <= current_time.tm_hour < 12:
            time_of_day = "morning"
        elif 12 <= current_time.tm_hour < 17:
            time_of_day = "afternoon"
        elif 17 <= current_time.tm_hour < 21:
            time_of_day = "evening"
        else:
            time_of_day = "night"
        
        # Adjust description based on weather and time
        if weather_condition in ['clear', 'clouds']:
            if time_of_day in ['morning', 'afternoon']:
                description = "with sunlight"
            elif time_of_day == 'evening':
                description = "during sunset"
            else:
                description = "under moonlight"
        elif weather_condition in ['rain', 'drizzle']:
            description = "with rainfall"
        elif weather_condition == 'thunderstorm':
            description = "with lightning"
        elif weather_condition == 'snow':
            description = "with snowfall"
        else:
            description = f"during {weather_condition} conditions"
        
        self.current_weather_conditions = f"{weather_condition.capitalize()}, {time_of_day}, {description}"
        
        return f"A photorealistic image of {weather_condition} weather during {time_of_day}, {description}"

    def apply_weather_overlay(self, img, weather_data):
        draw = ImageDraw.Draw(img, 'RGBA')
        font = ImageFont.truetype("arial.ttf", 20)

        temp_unit = self.temp_unit.get()
        temp = weather_data['main']['temp']
        feels_like = weather_data['main']['feels_like']
        
        if temp_unit == "F":
            temp = (temp * 9/5) + 32
            feels_like = (feels_like * 9/5) + 32

        weather_text = f"Temp: {temp:.1f}°{temp_unit}"
        weather_text += f"\nFeels like: {feels_like:.1f}°{temp_unit}"
        weather_text += f"\nWeather: {weather_data['weather'][0]['description']}"
        weather_text += f"\nHumidity: {weather_data['main']['humidity']}%"
        weather_text += f"\nWind: {weather_data['wind']['speed']:.2f} m/s"

        padding = 20
        line_spacing = 5
        
        # Calculate text dimensions
        lines = weather_text.split('\n')
        max_line_width = max(draw.textlength(line, font=font) for line in lines)
        text_height = sum(draw.textbbox((0, 0), line, font=font)[3] for line in lines) + line_spacing * (len(lines) - 1)
        
        is_span = self.wallpaper_style.get() == "span"
        img_width, img_height = img.size

        if is_span:
            left_monitor_width = img_width // 2
            right_padding = padding  # Use the same padding as the left side
            positions = {
                "left_top_left": (padding, padding),
                "left_top_right": (left_monitor_width - max_line_width - padding * 2 - right_padding, padding),
                "left_bottom_left": (padding, img_height - text_height - padding * 2),
                "left_bottom_right": (left_monitor_width - max_line_width - padding * 2 - right_padding, img_height - text_height - padding * 2),
                "right_top_left": (left_monitor_width + padding, padding),
                "right_top_right": (img_width - max_line_width - padding * 2 - right_padding, padding),
                "right_bottom_left": (left_monitor_width + padding, img_height - text_height - padding * 2),
                "right_bottom_right": (img_width - max_line_width - padding * 2 - right_padding, img_height - text_height - padding * 2)
            }
        else:
            positions = {
                "top_left": (padding, padding),
                "top_right": (img_width - max_line_width - padding * 2, padding),
                "bottom_left": (padding, img_height - text_height - padding * 2),
                "bottom_right": (img_width - max_line_width - padding * 2, img_height - text_height - padding * 2)
            }

        position = self.overlay_position.get()
        x, y = positions.get(position, (padding, padding))

        # Convert hex color to RGBA
        rgb_color = tuple(int(self.overlay_color.get()[i:i+2], 16) for i in (1, 3, 5))
        background_color = rgb_color + (self.overlay_opacity.get(),)

        # Add semi-transparent background
        background_bbox = [x, y, x + max_line_width + padding * 2, y + text_height + padding * 2]
        draw.rectangle(background_bbox, fill=background_color)

        # Apply drop shadow if enabled
        if self.use_drop_shadow.get():
            shadow_color = (0, 0, 0, 100)
            shadow_offset = 2
            for i, line in enumerate(lines):
                line_y = y + padding + i * (font.size + line_spacing) + shadow_offset
                draw.text((x + padding + shadow_offset, line_y), line, font=font, fill=shadow_color)

        # Draw text
        for i, line in enumerate(lines):
            line_y = y + padding + i * (font.size + line_spacing)
            draw.text((x + padding, line_y), line, font=font, fill=(255, 255, 255, 255))

        return img

    def apply_loaded_settings(self):
        # Update Start/Stop button
        if hasattr(self, 'start_stop_button'):
            self.start_stop_button.config(text="Stop" if getattr(self, 'is_running', False) else "Start")
        
        # Update Enhance checkbox
        if hasattr(self, 'enhance'):
            self.enhance.set(self.enhance.get())
        
        # Update Wallpaper Style dropdown
        if hasattr(self, 'wallpaper_style'):
            self.wallpaper_style.set(self.wallpaper_style.get())
        
        # Update Model dropdown
        if hasattr(self, 'model'):
            self.model.set(self.model.get())
        
        # Update Overlay Opacity scale
        if hasattr(self, 'overlay_opacity'):
            self.overlay_opacity.set(self.overlay_opacity.get())
        
        # Update Overlay Position dropdown
        if hasattr(self, 'overlay_position'):
            self.overlay_position.set(self.overlay_position.get())
        
        # Update Overlay Color button
        if hasattr(self, 'color_button'):
            self.color_button.config(bg=self.overlay_color.get())
        
        # Update Drop Shadow checkbox
        if hasattr(self, 'use_drop_shadow'):
            self.use_drop_shadow.set(self.use_drop_shadow.get())
        
        # Update Use Weather-Based Prompts checkbox
        if hasattr(self, 'use_weather'):
            self.use_weather.set(self.use_weather.get())
        
        # Update Overlay Weather Info checkbox
        if hasattr(self, 'overlay_weather'):
            self.overlay_weather.set(self.overlay_weather.get())
        
        # Update Temperature Unit radio buttons
        if hasattr(self, 'temp_unit'):
            self.temp_unit.set(self.temp_unit.get())
        
        # Update Always on Top
        if hasattr(self, 'always_on_top'):
            self.always_on_top.set(self.always_on_top.get())
        
        # Trigger necessary UI updates
        self.update_position_options()
        self.update_color_button()
        self.toggle_always_on_top()  # This line was mistakenly removed and is now added back

    def on_closing(self):
        self.stop_setter()
        if self.setter_thread and self.setter_thread.is_alive():
            self.setter_thread.join(timeout=5)
        self.save_settings()
        self.master.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = PollinationsBackgroundSetter(root)
    root.mainloop()