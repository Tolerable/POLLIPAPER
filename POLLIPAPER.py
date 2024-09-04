import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, colorchooser
import requests
from requests.exceptions import RequestException
import textwrap
from PIL import Image, ImageDraw, ImageFont, UnidentifiedImageError
import io
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
        master.geometry("650x450")

        self.enhance = tk.BooleanVar()
        self.wallpaper_style = tk.StringVar(value="fill")

        self.interval = max(300, 60 * 5)  # Minimum 5 minutes
        self.always_on_top = tk.BooleanVar(value=False)
        
        self.use_weather = tk.BooleanVar(value=False)
        self.overlay_weather = tk.BooleanVar(value=False)
        self.overlay_opacity = tk.IntVar(value=128)  # 0-255 for opacity
        self.overlay_position = tk.StringVar(value="top_right")
 
        self.use_drop_shadow = tk.BooleanVar(value=False)
        self.weather_api_key = tk.StringVar()
        self.weather_fetcher = weather_fetcher
        
        self.temp_unit = tk.StringVar(value="F")

        self.overlay_color = tk.StringVar(value="#000000")

        self.setup_ui()
        self.load_settings()
        self.update_position_options()
        self.prompt_history = deque(maxlen=20)
        self.load_history()
        
        self.is_running = False
        self.setter_thread = None
        self.current_request_id = None

        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.image_dir = "POLLINATIONS_BACKGROUNDS"
        os.makedirs(self.image_dir, exist_ok=True)

        self.style_dimensions = {
            "center": (1920, 1080),
            "tile": (1920, 1080),
            "stretch": (1920, 1080),
            "fit": (1920, 1080),
            "fill": (1920, 1080),
            "span": (1920, 540),
        }
        self.wallpaper_style.trace_add('write', self.update_position_options)
        
    def setup_ui(self):
        self.frame = ttk.Frame(self.master, padding="10")
        self.frame.pack(fill=tk.BOTH, expand=True)

        prompt_label = ttk.Label(self.frame, text="Prompt:")
        prompt_label.pack(fill=tk.X)

        prompt_frame = ttk.Frame(self.frame)
        prompt_frame.pack(fill=tk.BOTH, expand=True)

        self.prompt_entry = tk.Text(prompt_frame, wrap=tk.WORD, height=4)
        self.prompt_entry.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        prompt_scrollbar = ttk.Scrollbar(prompt_frame, orient="vertical", command=self.prompt_entry.yview)
        prompt_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.prompt_entry.config(yscrollcommand=prompt_scrollbar.set)

        history_label = ttk.Label(self.frame, text="History:")
        history_label.pack(fill=tk.X)
        self.history_var = tk.StringVar()
        self.history_dropdown = ttk.Combobox(self.frame, textvariable=self.history_var)
        self.history_dropdown.pack(fill=tk.X)
        self.history_dropdown.bind('<<ComboboxSelected>>', self.on_history_select)

        options_frame = ttk.LabelFrame(self.frame, text="Options", padding="5")
        options_frame.pack(fill=tk.X, pady=10)

        # First row
        first_row = ttk.Frame(options_frame)
        first_row.pack(fill=tk.X)

        self.start_stop_button = ttk.Button(first_row, text="Start", command=self.toggle_start_stop, width=10)
        self.start_stop_button.pack(side=tk.LEFT, padx=(0, 5))

        ttk.Checkbutton(first_row, text="Enhance", variable=self.enhance).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(first_row, text="Wallpaper Style:").pack(side=tk.LEFT)
        style_options = ["fill", "fit", "stretch", "tile", "center", "span"]
        ttk.Combobox(first_row, textvariable=self.wallpaper_style, values=style_options, width=10).pack(side=tk.LEFT)

        # Second row
        second_row = ttk.Frame(options_frame)
        second_row.pack(fill=tk.X, pady=(5, 0))

        ttk.Label(second_row, text="Overlay Opacity:").pack(side=tk.LEFT)
        ttk.Scale(second_row, from_=0, to=255, variable=self.overlay_opacity, orient=tk.HORIZONTAL, length=100).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(second_row, text="Position:").pack(side=tk.LEFT)
        self.position_dropdown = ttk.Combobox(second_row, textvariable=self.overlay_position, width=15)
        self.position_dropdown.pack(side=tk.LEFT, padx=(0, 10))
        self.update_position_options() 

        ttk.Label(second_row, text="Overlay Color:").pack(side=tk.LEFT)
        self.color_button = tk.Button(second_row, width=2, command=self.choose_overlay_color)
        self.color_button.pack(side=tk.LEFT, padx=(0, 10))
        self.update_color_button()

        ttk.Checkbutton(second_row, text="Drop Shadow", variable=self.use_drop_shadow).pack(side=tk.LEFT)

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

        self.current_prompt_label = ttk.Label(self.frame, text="Current Prompt:")
        self.current_prompt_label.pack(fill=tk.X, pady=(10, 0))
        
        self.current_prompt_display = tk.Text(self.frame, wrap=tk.WORD, height=3, state='disabled')
        self.current_prompt_display.pack(fill=tk.X)
    
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
        self.current_prompt_display.config(state='normal')
        self.current_prompt_display.delete(1.0, tk.END)
        self.current_prompt_display.insert(tk.END, prompt)
        self.current_prompt_display.config(state='disabled')

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
        if prompt and prompt not in self.prompt_history:
            self.prompt_history.appendleft(prompt)
            self.update_history_dropdown()
            self.save_history()

    def update_history_dropdown(self):
        self.history_dropdown['values'] = list(self.prompt_history)
        if self.prompt_history:
            self.history_var.set(self.prompt_history[0])

    def save_history(self):
        with open('prompt_history.json', 'w') as f:
            json.dump(list(self.prompt_history), f)

    def load_history(self):
        if os.path.exists('prompt_history.json'):
            with open('prompt_history.json', 'r') as f:
                self.prompt_history = deque(json.load(f), maxlen=20)
        self.update_history_dropdown()

    def start_setter(self):
        if self.use_weather.get():
            weather_data = self.weather_fetcher.fetch_weather_data()
            if weather_data:
                prompt = self.generate_weather_prompt(weather_data)
                self.update_current_prompt_display(prompt)
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
            self.add_to_history(prompt)
            self.update_current_prompt_display(prompt)

        self.is_running = True
        self.start_stop_button.config(text="Stop")
        self.current_request_id = random.randint(1, 1000000)
        
        self.setter_thread = threading.Thread(target=self.run_setter, args=(prompt, self.interval, self.current_request_id))
        self.setter_thread.start()

    def stop_setter(self):
        self.is_running = False
        self.start_stop_button.config(text="Start")
        self.current_request_id = None

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
            try:
                seed = random.randint(1, 1000000)
                enhance_param = "true" if self.enhance.get() else "false"
                style = self.wallpaper_style.get()
                
                width, height = self.get_image_dimensions(style)
                
                weather_data = None
                if self.use_weather.get() or self.overlay_weather.get():
                    weather_data = self.weather_fetcher.fetch_weather_data()
                    if self.use_weather.get() and weather_data:
                        prompt = self.generate_weather_prompt(weather_data)
                    elif self.use_weather.get():
                        print("Failed to fetch weather data. Using original prompt.")

                url = f"https://image.pollinations.ai/prompt/{prompt}, photographic_style?nologo=true&nofeed=true&enhance={enhance_param}&seed={seed}&width={width}&height={height}"

                print(f"Attempt {attempt + 1}: Fetching image with URL: {url}")
                
                print(f"Style: {style}, Dimensions: {width}x{height}")
                response = requests.get(url, timeout=45)
                
                if response.status_code == 200 and request_id == self.current_request_id:
                    print("Image fetched successfully.")
                    image = Image.open(io.BytesIO(response.content))
                    
                    image_path = os.path.join(self.image_dir, f"background_{int(time.time())}.png")
                    image.save(image_path, 'PNG')
                    print(f"Image saved to {image_path}")

                    if self.overlay_weather.get() and weather_data:
                        try:
                            image = self.apply_weather_overlay(image_path, weather_data)
                            image.save(image_path)
                            print(f"Weather overlay added to {image_path}")
                        except Exception as e:
                            print(f"Error applying weather overlay: {e}")

                    self.set_windows_background(image_path)
                    return
                else:
                    print(f"Failed to fetch image. Status code: {response.status_code}")
            
            except RequestException as e:
                print(f"Network error on attempt {attempt + 1}: {e}")
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
            "use_weather": self.use_weather.get(),
            "overlay_weather": self.overlay_weather.get(),
            "overlay_opacity": self.overlay_opacity.get(),
            "overlay_position": self.overlay_position.get(),
            "use_drop_shadow": self.use_drop_shadow.get(),
            "weather_api_key": self.weather_api_key.get(),
            "temp_unit": self.temp_unit.get()
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
                self.use_weather.set(settings.get("use_weather", False))
                self.overlay_weather.set(settings.get("overlay_weather", False))
                self.overlay_opacity.set(settings.get("overlay_opacity", 128))
                self.overlay_position.set(settings.get("overlay_position", "top_right"))
                self.use_drop_shadow.set(settings.get("use_drop_shadow", False))
                self.weather_api_key.set(settings.get("weather_api_key", ""))
                self.temp_unit.set(settings.get("temp_unit", "F"))
                
                if self.weather_api_key.get():
                    os.environ['OPENWEATHER_API_KEY'] = self.weather_api_key.get()
                    
                self.update_position_options()  # Add this line to update options after loading
                self.update_color_button()
                self.toggle_always_on_top()

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
        weather_condition = weather_data['weather'][0]['main']
        time_of_day = "day" if weather_data['dt'] < weather_data['sys']['sunset'] else "night"
        
        time_description = "sunny" if time_of_day == "day" else "with a visible moon"
        
        return f"A photographic image of {weather_condition.lower()} weather during {time_of_day}time, {time_description}"

    def apply_weather_overlay(self, image_path, weather_data):
        img = Image.open(image_path)
        draw = ImageDraw.Draw(img, 'RGBA')
        font = ImageFont.truetype("arial.ttf", 20)

        temp_unit = self.temp_unit.get()
        temp = weather_data['main']['temp']
        feels_like = weather_data['main']['feels_like']
        
        if temp_unit == "F":
            temp = (temp * 9/5) + 32
            feels_like = (feels_like * 9/5) + 32

        weather_text = (f"Temp: {temp:.1f}°{temp_unit}\n"
                        f"Feels like: {feels_like:.1f}°{temp_unit}\n"
                        f"Weather: {weather_data['weather'][0]['description']}\n"
                        f"Humidity: {weather_data['main']['humidity']}%\n"
                        f"Wind: {weather_data['wind']['speed']} m/s")

        wrapped_text = textwrap.wrap(weather_text, width=23)
        
        padding = 20
        text_width, text_height = draw.multiline_textbbox((0, 0), "\n".join(wrapped_text), font=font)[2:]
        
        is_span = self.wallpaper_style.get() == "span"
        img_width, img_height = img.size

        if is_span:
            left_monitor_width = img_width // 2
            right_monitor_width = img_width - left_monitor_width
            positions = {
                "left_top_left": (padding, padding),
                "left_top_right": (left_monitor_width - text_width - padding, padding),
                "left_bottom_left": (padding, img_height - text_height - padding),
                "left_bottom_right": (left_monitor_width - text_width - padding, img_height - text_height - padding),
                "right_top_left": (left_monitor_width + padding, padding),
                "right_top_right": (img_width - text_width - padding, padding),
                "right_bottom_left": (left_monitor_width + padding, img_height - text_height - padding),
                "right_bottom_right": (img_width - text_width - padding, img_height - text_height - padding)
            }
        else:
            positions = {
                "top_left": (padding, padding),
                "top_right": (img_width - text_width - padding, padding),
                "bottom_left": (padding, img_height - text_height - padding),
                "bottom_right": (img_width - text_width - padding, img_height - text_height - padding)
            }

        position = self.overlay_position.get()
        x, y = positions.get(position, (padding, padding))  # Default to top-left if position is not found

        # Convert hex color to RGBA
        rgb_color = tuple(int(self.overlay_color.get()[i:i+2], 16) for i in (1, 3, 5))
        background_color = rgb_color + (self.overlay_opacity.get(),)

        # Add semi-transparent background
        background_bbox = [x, y, x + text_width + padding, y + text_height + padding]
        draw.rectangle(background_bbox, fill=background_color)

        # Apply drop shadow if enabled
        if self.use_drop_shadow.get():
            shadow_color = (0, 0, 0, 100)
            shadow_offset = 2
            draw.multiline_text((x + shadow_offset, y + shadow_offset), "\n".join(wrapped_text), font=font, fill=shadow_color)

        # Draw text
        draw.multiline_text((x, y), "\n".join(wrapped_text), font=font, fill=(255, 255, 255, 255))

        return img

    def save_settings(self):
        settings = {
            "enhance": self.enhance.get(),
            "wallpaper_style": self.wallpaper_style.get(),
            "interval": self.interval,
            "overlay_color": self.overlay_color.get(),
            "always_on_top": self.always_on_top.get()
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
                self.overlay_color.set(settings.get("overlay_color", "#000000"))
                                
        self.update_color_button()
        self.toggle_always_on_top()

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