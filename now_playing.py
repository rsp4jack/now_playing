#!/usr/bin/env python
# title             : now_playing.py
# description       : Now Playing is an OBS script that will update a Text Source
#                   : with the current song that Media Player are playing. Only for Windows OS
# author            : Etuldan(Orgin)
#                   : Creepercdn(Fork)
# date              : 2020 12 12
# version           : 0.2
# usage             : python now_playing.py
# dependencies      : - Python 3.6 (https://www.python.org/)
#                   :   - pywin32 (https://github.com/mhammond/pywin32/releases)
#                   : - Windows Vista+
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

import ctypes
import ctypes.wintypes
import site
from types import LambdaType
from typing import AnyStr, Sequence

import win32api
import win32con
import win32gui
import win32process

import obspython as obs

import os

site.main()

enabled = True
check_frequency = 1000  # ms
display_text = '%artist - %title'
debug_mode = True

source_name = ''
calcingfunc = {}

class Capture(object):
    def __init__(self, id: AnyStr, display_name: AnyStr, state: bool) -> None:
        self.id = id
        self.display_name = display_name
        self.state = state
    
    def capture(self, title: AnyStr, process: AnyStr) -> Sequence[str]:
        """Capture the song info

        Args:
            title (AnyStr): [description]
            process (AnyStr): [description]

        Returns:
            Sequence[str]: 0:Artist, 1:Song
        """
        # print((lambda x,y: ((x[0:x.find('-')-1], x[x.find('-')+2:x.rfind('-')-1]) if (y.lower().endswith('vlc.exe') and captures.vlc.state and "-" in x) else (None, None)))(title, process))
        return calcingfunc.get(self.id, lambda x,y: ('Underdefined', 'ERR'))(title, process)


# We defined some factory presets here for you.
# BEGIN PRESETS DEFINE

# BEGIN SPECIAL CAPTURES
def foobar2000capture(title: AnyStr, process: AnyStr) -> Sequence[str]:
    artist = ''
    song = ''
    if foobar2000.state and process.lower().endswith('foobar2000.exe'):
        if ("-" not in title) and (title.find('[foobar2000]') != -1):
            song = title[:title.rfind(" [foobar2000]")-1]
        elif "-" in title:
            artist = title[0:title.find("-")-1]
            song = title[title.find("]")+2:title.rfind(" [foobar2000]")-1]
    return (artist, song)

def serato_capture(url: str):
    # Thanks DachsbauTV for this feature! (#2)
    # TODO: Add Serato capture (#2)
    pass

# END SPECIAL CAPTURES

spotify = Capture('spotify', 'Spotify', True)
calcingfunc['spotify'] = lambda x,y: ((x[0:x.find('-')-1], x[x.find('-')+2:]) if (y.lower().endswith('spotify.exe') and spotify.state and "-" in x) else (None, None))
vlc = Capture('vlc', "VLC", True)
calcingfunc['vlc'] = lambda x,y: ((x[0:x.find('-')-1], x[x.find('-')+2:x.rfind('-')-1]) if (y.lower().endswith('vlc.exe') and vlc.state and "-" in x) else (None, None))
yt_firefox = Capture('yt_firefox', "YouTube for Firefox", True)
calcingfunc['yt_firefox'] = lambda x,y: ((x[0:x.find('-')-1], x[x.find('-')+2:x.rfind('-')-1]) if (y.lower().endswith('firefox.exe') and yt_firefox.state and "- YouTube" in x) else (None, None))
yt_chrome = Capture('yt_chrome', 'YouTube for Chrome', True)
calcingfunc['yt_chrome'] = lambda x,y: ((x[0:x.find('-')-1], x[x.find('-')+2:x.rfind('-')-1]) if (y.lower().endswith('chrome.exe') and yt_chrome.state and "- YouTube" in x) else (None, None))

foobar2000 = Capture('foobar2000', 'Foobar2000', True)
calcingfunc['foobar2000'] = foobar2000capture
necloud = Capture('necloud', 'Netease Cloud Music', True)
calcingfunc['necloud'] = lambda x,y: ((x[x.find("-")+2:], x[0:x.find("-")-1]) if (y.lower().endswith('cloudmusic.exe') and necloud.state and "-" in x) else (None, None))
aimp = Capture('aimp', 'AIMP', True)
calcingfunc['aimp'] = lambda x,y: ((x[0:x.find('-')-1], x[x.find('-')+2:]) if (y.lower().endswith('aimp.exe') and aimp.state and "-" in x) else (None, None))

# END PRESETS DEFINE

def debug(*args, sep: str = ' ', end: str = '\n', flush: bool = False) -> None:
    """A debug info printer

    Args:
        sep (str, optional): [description]. Defaults to ' '.
        end (str, optional): [description]. Defaults to '\n'.
        flush (bool, optional): [description]. Defaults to False.
    """
    if debug_mode:
        print(*args, sep=sep, end=end, flush=flush)

def IsWindowVisibleOnScreen(hwnd):
    def IsWindowCloaked(hwnd):
        DWMWA_CLOAKED = 14
        cloaked = ctypes.wintypes.DWORD()
        ctypes.windll.dwmapi.DwmGetWindowAttribute(hwnd, ctypes.wintypes.DWORD(
            DWMWA_CLOAKED), ctypes.byref(cloaked), ctypes.sizeof(cloaked))
        return cloaked.value
    return ctypes.windll.user32.IsWindowVisible(hwnd) and (not IsWindowCloaked(hwnd))

def callsmtc():
    # TODO: SMTC.py supports
    pass


def script_defaults(settings):
    debug("Calling defaults")

    obs.obs_data_set_default_bool(settings, "enabled", enabled)
    obs.obs_data_set_default_int(settings, "check_frequency", check_frequency)
    obs.obs_data_set_default_string(settings, "display_text", display_text)
    obs.obs_data_set_default_string(settings, "source_name", source_name)
    for i in filter(lambda x: not x.startswith('_'), calcingfunc.keys()):
        obs.obs_data_set_default_bool(settings, globals()[i].id, globals()[i].state)


def script_description():
    debug("Calling description")

    return \
        """
        <h1>Now Playing by Creepercdn</h1>
        <hr/>
        Display current song as a text on your screen.
        <br/>
        Available placeholders:
        <br/>
        <code>%artist</code>, <code>%title</code>
        <br/>
        GitHub:
        <a href="https://github.com/Creepercdn/now_playing">Creepercdn/now_playing</a>
        <hr/>
        """


def script_load(_):
    debug("[CS] Loaded script.")


def script_properties():
    debug("[CS] Loaded properties.")

    props = obs.obs_properties_create()
    obs.obs_properties_add_bool(props, "enabled", "Enabled")
    obs.obs_properties_add_bool(props, "debug_mode", "Debug Mode")
    obs.obs_properties_add_int(
        props, "check_frequency", "Check frequency", 150, 10000, 100)
    obs.obs_properties_add_text(
        props, "display_text", "Display text", obs.OBS_TEXT_DEFAULT)
    for i in filter(lambda x: not x.startswith('_'), calcingfunc.keys()):
        obs.obs_properties_add_bool(props, globals()[i].id, globals()[i].display_name)

    p = obs.obs_properties_add_list(
        props, "source_name", "Text source",
        obs.OBS_COMBO_TYPE_EDITABLE, obs.OBS_COMBO_FORMAT_STRING)

    sources = obs.obs_enum_sources()
    if sources:
        for source in sources:
            source_id = obs.obs_source_get_unversioned_id(source)
            if source_id in ("text_gdiplus", "text_ft2_source"):
                name = obs.obs_source_get_name(source)
                obs.obs_property_list_add_string(p, name, name)
    obs.source_list_release(sources)

    return props


def script_save(settings):
    debug("[CS] Saved properties.")

    script_update(settings)


def script_unload():
    debug("[CS] Unloaded script.")

    obs.timer_remove(get_song_info)


def script_update(settings):
    global enabled
    global display_text
    global check_frequency
    global source_name
    global debug_mode
    debug("[CS] Updated properties.")

    if obs.obs_data_get_bool(settings, "enabled"):
        if not enabled:
            debug("[CS] Enabled song timer.")
            enabled = True
            obs.timer_add(get_song_info, check_frequency)
    else:
        if enabled:
            debug("[CS] Disabled song timer.")
            enabled = False
            obs.timer_remove(get_song_info)

    debug_mode = obs.obs_data_get_bool(settings, "debug_mode")
    display_text = obs.obs_data_get_string(settings, "display_text")
    source_name = obs.obs_data_get_string(settings, "source_name")
    check_frequency = obs.obs_data_get_int(settings, "check_frequency")
    for i in filter(lambda x: not x.startswith('_'), calcingfunc.keys()):
        globals()[i].state = obs.obs_data_get_bool(settings, globals()[i].id)



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
    debug("[CS] Now Playing : " + artist + " / " + song)


def get_song_info():

    def enumHandler(hwnd, result):

        _, procpid = win32process.GetWindowThreadProcessId(hwnd)
        try:
            if not IsWindowVisibleOnScreen(hwnd):
                return
            mypyproc = win32api.OpenProcess(
                win32con.PROCESS_ALL_ACCESS, False, procpid)
            exe = win32process.GetModuleFileNameEx(mypyproc, 0)
            title = win32gui.GetWindowText(hwnd)
            for i in filter(lambda x: not x.startswith('_'), calcingfunc.keys()):
                res = globals()[i].capture(title, exe)
                
                # 0:Arist, 1:Song
                artist = ''
                song = ''
                if res[0]:
                    artist = os.path.splitext(res[0])[0]
                if res[1]:
                    song = os.path.splitext(res[1])[0]
                if any([artist, song]):
                    result.append([artist, song])
        except:
            return
        return

    result = []
    win32gui.EnumWindows(enumHandler, result)
    try:
        update_song(result[0][0], result[0][1])
    except:
        update_song()
