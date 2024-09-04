
# POLLIPAPER

POLLIPAPER is a dynamic wallpaper application that uses the Pollinations AI image generation service to create and set unique desktop backgrounds. It offers features like weather-based prompts, weather information overlay, and customizable settings.

## Features

- Generate AI-powered wallpapers using custom prompts
- Automatic wallpaper changing at set intervals
- Weather-based prompt generation
- Weather information overlay on wallpapers
- Customizable wallpaper styles (fill, fit, stretch, tile, center, span)
- Enhance option for higher quality images
- History of used prompts
- Adjustable overlay opacity and position
- Drop shadow option for weather overlay
- Temperature unit selection (Fahrenheit/Celsius)
- Always-on-top window option

## Requirements

- Python 3.6+
- tkinter
- Pillow
- requests
- comtypes (for Windows)

## Installation

1. Clone the repository:
   ```
   git clone https://github.com/yourusername/pollipaper.git
   cd pollipaper
   ```

2. Install the required packages:
   ```
   pip install -r requirements.txt
   ```

3. Run the application:
   ```
   python pollipaper.py
   ```

## Usage

1. Enter a prompt in the text box or select one from the history dropdown.
2. Click the "Start" button to begin generating and setting wallpapers.
3. Adjust settings as desired using the Options menu.
4. To use weather-based prompts or overlay weather information, set your OpenWeatherMap API key in the Options menu.

## Configuration

You can customize various aspects of POLLIPAPER:

- **Enhance**: Toggle to generate higher quality images (may increase generation time)
- **Wallpaper Style**: Choose how the wallpaper is displayed on your desktop
- **Interval**: Set how often the wallpaper changes
- **Weather Options**: Enable weather-based prompts or weather information overlay
- **Overlay Settings**: Adjust opacity, position, and color of the weather overlay
- **Temperature Unit**: Choose between Fahrenheit and Celsius

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- [Pollinations AI](https://pollinations.ai/) for providing the image generation service
- [OpenWeatherMap](https://openweathermap.org/) for weather data

## Disclaimer

This application is for personal use only. Please ensure you comply with Pollinations AI's terms of service when using this application.
