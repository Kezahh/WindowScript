"""Main WINDOWSCRIPT."""

###############################################################################
# Copyright (C) Kerryn Spanhel - All Rights Reserved
# Unauthorized copying of this file is at user's own risk.
# Written by Kerryn Spanhel, 2020
###############################################################################

import sys
import ctypes
from ctypes import windll, byref, Structure, wintypes, sizeof


# Define debug output path
OUTPUT_PATH = os.path.join(os.environ["USERPROFILE"], r"Desktop\output.txt")

# Define Windows ENUM constant
DWMA_EXTENDED_FRAME_BOUNDS = 9


class RECT(Structure):
    """RECT structure needed for user32.dll function returns."""
    
    _fields_ = [
        # Order of fields is important. Must be like this.
        ('left', wintypes.ULONG),
        ('top', wintypes.ULONG),
        ('right', wintypes.ULONG),
        ('bottom', wintypes.ULONG)
    ]
    
    @property
    def x(self):
        return self._convert_ulong_to_int(self.left)
    @property
    def y(self):
        return self._convert_ulong_to_int(self.top)
    @property
    def w(self):
        return (self._convert_ulong_to_int(self.right) - self._convert_ulong_to_int(self.left))
    @property
    def h(self):
        return (self._convert_ulong_to_int(self.bottom) - self._convert_ulong_to_int(self.top))
    
    @staticmethod
    def _convert_ulong_to_int(x):
        return x if x < (1<<31) else (x - (1<<32))


def print_debug(message)
    """Print message to debug file.
    
    Script needs to run with no console. Debug file is used to see debug 
    messages.
    """
    try:
        output_file.write(message + "\n")
    except NameError:
        pass


def main():
    if len(sys.argv) > 1:
        try:
            POSITION_INDEX = int(sys.argv[1])
        except ValueError:
            POSITION_INDEX = 0
    else:
        POSITION_INDEX = 0

    output_file.write(str(POSITION_INDEX))
    if (POSITION_INDEX in range(0,3)):
        x_left = 1706*POSITION_INDEX
        target_width = 1706
        target_height = 1400
    elif (POSITION_INDEX == 3):
        x_left = 1280
        target_width = 2560
        target_height = 1400

    y_top = 0

    window_handle = windll.user32.GetForegroundWindow()

    handles = []
    i = -1
    while (True):
        i += 1
        handles.append(window_handle)
        window_handle = windll.user32.GetWindow(window_handle, 2)
        parent_handle = window_handle
        next_handle = windll.user32.GetParent(parent_handle)
        while (next_handle != 0):
            parent_handle = next_handle
            next_handle = windll.user32.GetParent(parent_handle)
        process_id = ctypes.c_ulong()
        result = windll.user32.GetWindowThreadProcessId(parent_handle, byref(process_id))
        
        window_text_length = windll.user32.GetWindowTextLengthA(window_handle)
        window_text = ctypes.create_string_buffer(window_text_length + 1)
        windll.user32.GetWindowTextA(window_handle, byref(window_text), window_text_length + 1)
        
        try:
            actual_text = str(window_text.value)
        except UnicodeEncodeError:
            actual_text = ""
        
        if (windll.user32.IsWindowVisible(window_handle) == 1):
            window_handle = parent_handle
            break
        
        
        
    window_rect = RECT()
    client_rect = RECT()
    frame_rect = RECT()

    windll.user32.GetWindowRect(window_handle, byref(window_rect))
    windll.user32.GetClientRect(window_handle, byref(client_rect))
    windll.dwmapi.DwmGetWindowAttribute(window_handle, DWMA_EXTENDED_FRAME_BOUNDS, byref(frame_rect), sizeof(frame_rect))


    border_width = (window_rect.w - frame_rect.w)
    border_height = (window_rect.h - frame_rect.h)

    #import pdb; pdb.set_trace()

    window_new_x = x_left + (window_rect.x - frame_rect.x)
    window_new_width = target_width + border_width
    window_new_height = target_height + border_height

    windll.user32.MoveWindow(
        window_handle, window_new_x, 0,
        window_new_width, window_new_height, True
    )

if __name__ == "__main__":
    with open(OUTPUT_PATH, 'w') as output_file:
        main()