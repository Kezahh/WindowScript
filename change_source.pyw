"""Main WINDOWSCRIPT."""

###############################################################################
# Copyright (C) Kerryn Spanhel - All Rights Reserved
# Unauthorized copying of this file is at user's own risk.
# Written by Kerryn Spanhel, 2020
###############################################################################

import os
import sys
import ctypes
from ctypes import windll, byref, Structure, WinError, POINTER, WINFUNCTYPE
from ctypes.wintypes import BOOL, HMONITOR, HDC, RECT, LPARAM, DWORD, BYTE, WCHAR, HANDLE
import time


# Define debug output path
OUTPUT_PATH = os.path.join(os.environ["USERPROFILE"], r"Desktop\output.txt")

# Define monitor constants
DISPLAYPORT = 0xF0F
USBC = 0xF1B

# Define Windows ENUM constant
_MONITORENUMPROC = WINFUNCTYPE(BOOL, HMONITOR, HDC, POINTER(RECT), LPARAM)

class PHYSICAL_MONITOR(Structure):
    _fields_ = [
        ('handle', HANDLE),
        ('description', WCHAR * 128)
    ]

def _iter_physical_monitors(close_handles=True):
    """Iterates physical monitors.

    The handles are closed automatically whenever the iterator is advanced.
    This means that the iterator should always be fully exhausted!

    If you want to keep handles e.g. because you need to store all of them and
    use them later, set `close_handles` to False and close them manually."""

    def callback(hmonitor, hdc, lprect, lparam):
        monitors.append(HMONITOR(hmonitor))
        return True

    monitors = []
    if not windll.user32.EnumDisplayMonitors(None, None, _MONITORENUMPROC(callback), None):
        raise WinError('EnumDisplayMonitors failed')

    for monitor in monitors:
        # Get physical monitor count
        count = DWORD()
        if not windll.dxva2.GetNumberOfPhysicalMonitorsFromHMONITOR(monitor, byref(count)):
            raise WinError()
        # Get physical monitor handles
        physical_array = (PHYSICAL_MONITOR * count.value)()
        if not windll.dxva2.GetPhysicalMonitorsFromHMONITOR(monitor, count.value, physical_array):
            raise WinError()
        for physical in physical_array:
            yield physical.handle
            if close_handles:
                if not windll.dxva2.DestroyPhysicalMonitor(physical.handle):
                    raise WinError()
                    
def set_vcp_feature(monitor, code, value):
    """Sends a DDC command to the specified monitor.
    """
    if not windll.dxva2.SetVCPFeature(HANDLE(monitor), BYTE(code), DWORD(value)):
        raise WinError()
        
def get_vcp_feature(monitor, code):
    """Sends a DDC command to specific monitor.
    """
    current_value = DWORD()
    max_value = DWORD()
    if not windll.dxva2.GetVCPFeatureAndVCPFeatureReply(
        HANDLE(monitor), BYTE(code), None, byref(current_value), byref(max_value)):
        raise WinError()
    
    return current_value.value, max_value.value


# Switch to HDMI, wait for the user to press return and then back to DP
for handle in _iter_physical_monitors():
    #set_vcp_feature(handle, 0x60, 0x0F) #Change input to DisplayPort
    #set_vcp_feature(handle, 0x60, 0x1B) #Change input to USB-C
    current_source = get_vcp_feature(handle, 0x60)[0]
    if current_source == DISPLAYPORT:
        target_source = USBC
    else:
        target_source = DISPLAYPORT
    
    set_vcp_feature(handle, 0x60, target_source)
    
    #need to sleep then set the USB to USB1. This helps when screen is asleep.
    time.sleep(1)
    set_vcp_feature(handle, 0xE7, 0)