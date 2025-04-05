# Windows-Automation-Toolkit

This Python library provides a set of tools for automating tasks on Windows operating systems. It allows you to control mouse and keyboard inputs, as well as interact with the screen by getting and setting pixel colors.

## Features

* **Mouse Control:**
    * Get and set the current cursor position.
    * Move the cursor relatively.
    * Simulate left, right, and middle mouse button clicks (press and release).
    * Simulate mouse button down and up events.
    * Scroll the mouse wheel.
* **Keyboard Control:**
    * Press and release individual keys.
    * Type text.
    * Wait for a specific key to be pressed.
    * Check if a key is currently pressed.
* **Screen Interaction:**
    * Get the RGB color of a pixel at a specific coordinate or the current mouse position.
    * Set the RGB color of a pixel at a specific coordinate.

## Installation

1.  Make sure you have Python installed on your Windows system.
2.  Install the required dependencies using pip:

    ```bash
    pip install pynput
    pip install pywin32
    pip install keyboard
    ```

## Usage

Here are some basic examples of how to use the `PC` class from the library:

```python
import time
from Control import PC
from pynput.keyboard import Key

# Initialize the PC automation object
pc = PC()

# Mouse Control Examples
current_pos = pc.get_pos()
print(f"Current mouse position: {current_pos}")

pc.set_pos(100, 100)  # Move mouse to coordinates (100, 100)
pc.click()  # Left click at the current position
pc.click(500, 300, button="right")  # Right click at coordinates (500, 300)
pc.scroll(1)  # Scroll up
pc.scroll(-1) # Scroll down

# Keyboard Control Examples
pc.write("Hello, World!")  # Type the text "Hello, World!"
pc.press("ENTER")  # Press the Enter key
pc.key_down("SHIFT")
pc.press("A")  # Types a capital 'A'
pc.key_up("SHIFT")

print("Press the 'Q' key to continue...")
pc.wait("q", beep=500) # Wait for 'q' key press with a beep sound

# Screen Interaction Examples
pixel_color = pc.get_pixel(200, 200)
print(f"Color at (200, 200): {pixel_color}")

# Set a pixel to red (RGB: 255, 0, 0)
pc.set_pixel(300, 300, 255, 0, 0)
