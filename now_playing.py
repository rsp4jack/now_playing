#!/usr/bin/env python
# title             : now_playing.py
# description       : Now Playing is an OBS script that will update a Text Source
#                   : with the current song that Media Player are playing. Only for Windows OS
# author            : Etuldan
# date              : 2019 03 30
# version           : 0.1
# usage             : python now_playing.py
# dependencies      : - Python 3.6 (https://www.python.org/)
#                   :   - pywin32 (https://github.com/mhammond/pywin32/releases)
# notes             : Follow this step for this script to work:
#                   : Python:
#                   :   1. Install python (v3.6 and 64 bits, this is important)
#                   :   2. Install pywin32 (not with pip, but with installer)
#                   : OBS:
#                   :   1. Create a GDI+ Text Source with the name of your choice
#                   :   2. Go to Tools › Scripts
#                   :   3. Click the "+" button and add this script
#                   :   5. Set the same source name as the one you just created
#                   :   6. Check "Enable"
#                   :   7. Click the "Python Settings" rab
#                   :   8. Select your python install path
#                   :
# python_version    : 3.6+
# ==============================================================================

import site
import obspython as obs
import win32gui
import win32process
import win32api
import win32con
import ctypes
import ctypes.wintypes


working = True
enabled = True
check_frequency = 1000
display_text = '%artist - %title'
debug_mode = True

source_name = ''

spotify = True
vlc = True
yt_firefox = True
yt_chrome = True
foobar2000 = True
necloud = True
aimp = True
potplayer = True


def IsWindowVisibleOnScreen(hwnd):
    def IsWindowCloaked(hwnd):
        DWMWA_CLOAKED = 14
        cloaked = ctypes.wintypes.DWORD()
        ctypes.windll.dwmapi.DwmGetWindowAttribute(hwnd, ctypes.wintypes.DWORD(
            DWMWA_CLOAKED), ctypes.byref(cloaked), ctypes.sizeof(cloaked))
        return cloaked.value
    return ctypes.windll.user32.IsWindowVisible(hwnd) and (not IsWindowCloaked(hwnd))


def script_defaults(settings):
    if debug_mode:
        print("Calling defaults")

    obs.obs_data_set_default_bool(settings, "enabled", enabled)
    obs.obs_data_set_default_int(settings, "check_frequency", check_frequency)
    obs.obs_data_set_default_string(settings, "display_text", display_text)
    obs.obs_data_set_default_string(settings, "source_name", source_name)
    obs.obs_data_set_default_bool(settings, "spotify", spotify)
    obs.obs_data_set_default_bool(settings, "vlc", vlc)
    obs.obs_data_set_default_bool(settings, "yt_firefox", yt_firefox)
    obs.obs_data_set_default_bool(settings, "yt_chrome", yt_chrome)
    obs.obs_data_set_default_bool(settings, "foobar2000", foobar2000)
    obs.obs_data_set_default_bool(settings, "necloud", necloud)
    obs.obs_data_set_default_bool(settings, "aimp", aimp)
    obs.obs_data_set_default_bool(settings, "potplayer", potplayer)


def script_description():
    if debug_mode:
        print("Calling description")

    return "<b>Music Now Playing</b>" + \
        "<hr>" + \
        "Display current song as a text on your screen." + \
        "<br/>" + \
        "Available placeholders: " + \
        "<br/>" + \
        "<code>%artist</code>, <code>%title</code>" + \
        "<hr>"


def script_load(settings):
    if debug_mode:
        print("[CS] Loaded script.")


def script_properties():
    if debug_mode:
        print("[CS] Loaded properties.")

    props = obs.obs_properties_create()
    obs.obs_properties_add_bool(props, "enabled", "Enabled")
    obs.obs_properties_add_bool(props, "debug_mode", "Debug Mode")
    obs.obs_properties_add_int(
        props, "check_frequency", "Check frequency", 150, 10000, 100)
    obs.obs_properties_add_text(
        props, "display_text", "Display text", obs.OBS_TEXT_DEFAULT)
    obs.obs_properties_add_bool(props, "spotify", "Spotify")
    obs.obs_properties_add_bool(props, "vlc", "VLC")
    obs.obs_properties_add_bool(props, "yt_firefox", "Youtube for Firefox")
    obs.obs_properties_add_bool(props, "yt_chrome", "Youtube for Chrome")
    obs.obs_properties_add_bool(props, "foobar2000", "Foobar2000")
    obs.obs_properties_add_bool(props, "necloud", "Netease Cloud Music")
    obs.obs_properties_add_bool(props, 'aimp', 'AIMP')
    obs.obs_properties_add_bool(props, 'potplayer', 'PotPlayer(Only File name)')
    obs.obs_properties_add_text(
        props, "source_name", "Text source", obs.OBS_TEXT_DEFAULT)
    return props


def script_save(settings):
    if debug_mode:
        print("[CS] Saved properties.")

    script_update(settings)


def script_unload():
    if debug_mode:
        print("[CS] Unloaded script.")

    obs.timer_remove(get_song_info)


def script_update(settings):
    global debug_mode
    if debug_mode:
        print("[CS] Updated properties.")

    global enabled
    global display_text
    global check_frequency
    global source_name
    global spotify
    global vlc
    global yt_firefox
    global yt_chrome
    global foobar2000
    global necloud
    global aimp
    global potplayer

    if obs.obs_data_get_bool(settings, "enabled") is True:
        if (not enabled):
            if debug_mode:
                print("[CS] Enabled song timer.")

        enabled = True
        obs.timer_add(get_song_info, check_frequency)
    else:
        if (enabled):
            if debug_mode:
                print("[CS] Disabled song timer.")

        enabled = False
        obs.timer_remove(get_song_info)

    debug_mode = obs.obs_data_get_bool(settings, "debug_mode")
    display_text = obs.obs_data_get_string(settings, "display_text")
    source_name = obs.obs_data_get_string(settings, "source_name")
    check_frequency = obs.obs_data_get_int(settings, "check_frequency")
    spotify = obs.obs_data_get_bool(settings, "spotify")
    vlc = obs.obs_data_get_bool(settings, "vlc")
    yt_firefox = obs.obs_data_get_bool(settings, "yt_firefox")
    yt_chrome = obs.obs_data_get_bool(settings, "yt_chrome")
    foobar2000 = obs.obs_data_get_bool(settings, "foobar2000")
    necloud = obs.obs_data_get_bool(settings, "necloud")
    aimp = obs.obs_data_get_bool(settings, "aimp")
    potplayer = obs.obs_data_get_bool(settings, "potplayer")


def update_song(artist="", song=""):

    now_playing = ""
    if(artist != "" or song != ""):
        now_playing = display_text.replace(
            '%artist', artist).replace('%title', song)

    settings = obs.obs_data_create()
    obs.obs_data_set_string(settings, "text", now_playing)
    source = obs.obs_get_source_by_name(source_name)
    obs.obs_source_update(source, settings)
    obs.obs_data_release(settings)
    obs.obs_source_release(source)
    if debug_mode:
        print("[CS] Now Playing : " + artist + " / " + song)


def get_song_info():

    def enumHandler(hwnd, result):

        _, procpid = win32process.GetWindowThreadProcessId(hwnd)
        try:
            if not IsWindowVisibleOnScreen(hwnd):
                return
            mypyproc = win32api.OpenProcess(
                win32con.PROCESS_ALL_ACCESS, False, procpid)
            exe = win32process.GetModuleFileNameEx(mypyproc, 0)
            if spotify and exe.endswith("Spotify.exe"):
                title = win32gui.GetWindowText(hwnd)
                if("-" in title):
                    artist = title[0:title.find("-")-1]
                    song = title[title.find("-")+2:]
                    result.append([artist, song])
                    return
            if vlc and exe.endswith("vlc.exe"):
                title = win32gui.GetWindowText(hwnd)
                if("-" in title):
                    artist = title[0:title.find("-")-1]
                    song = title[title.find("-")+2:title.rfind("-")-1]
                    result.append([artist, song])
                    return
            if yt_firefox and exe.endswith("firefox.exe"):
                title = win32gui.GetWindowText(hwnd)
                if("- YouTube" in title):
                    artist = title[0:title.find("-")-1]
                    song = title[title.find("-")+2:title.rfind("-")-1]
                    result.append([artist, song])
                    return
            if yt_chrome and exe.endswith("chrome.exe"):
                title = win32gui.GetWindowText(hwnd)
                if("- YouTube" in title):
                    artist = title[0:title.find("-")-1]
                    song = title[title.find("-")+2:title.rfind("-")-1]
                    result.append([artist, song])
                    return
            if foobar2000 and exe.endswith("foobar2000.exe"):
                title = win32gui.GetWindowText(hwnd)
                if("-" in title):
                    artist = title[0:title.find("-")-1]
                    song = title[title.find("]") +
                                 2:title.rfind(" [foobar2000]")-1]
                    result.append([artist, song])
                    return
            if necloud and exe.endswith("cloudmusic.exe"):
                title = win32gui.GetWindowText(hwnd)
                if("-" in title):
                    song = title[0:title.find("-")-1]
                    artist = title[title.find("-")+2:]
                    result.append([artist, song])
                    return
            if necloud and exe.endswith("cloudmusic.exe"):
                title = win32gui.GetWindowText(hwnd)
                if("-" in title):
                    song = title[0:title.find("-")-1]
                    artist = title[title.find("-")+2:]
                    result.append([artist, song])
                    return
            if aimp and exe.endswith("AIMP.exe"):
                title = win32gui.GetWindowText(hwnd)
                if("-" in title):
                    artist = title[0:title.find("-")-1]
                    song = title[title.find("-")+2:]
                    result.append([artist, song])
                    return
            if potplayer and exe.endswith('PotPlayerMini.exe'):
                title = win32gui.GetWindowText(hwnd)
                if ('-' in title):
                    artist = 'Unknown'
                    song = title[0:title.rfind(' PotPlayer')-2]
                    result.append([artist, song])
                    return
            
        except Exception:
            return
        return

    result = []
    win32gui.EnumWindows(enumHandler, result)
    try:
        update_song(result[0][0], result[0][1])
    except:
        update_song()
