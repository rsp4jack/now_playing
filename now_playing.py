#!/usr/bin/env python
# title             : now_playing.py
# description       : Now Playing is an OBS script that will update a Text Source
#                   : with the current song that a media player are playing. 
#                   : Only for Windows now
# author            : Etuldan
#                   : rsp4jack
# version           : 0.2
# usage             : python now_playing.py
# dependencies      : - Python 3.7+ (https://www.python.org/)
#                   :   - pywin32 (https://github.com/mhammond/pywin32/releases)
#                   : - Windows Vista+
# notes             : 
# python_version    : 3.7+
# ==============================================================================

def script_description():
    log.debug("script_description()")

    return \
        """
        <h1>Now Playing</h1>
        <hr/>
        Display current playing song as text.
        <br/>
        Available placeholders:
        <br/>
        <code>%artist</code>, <code>%title</code>
        <br/>
        <a href="https://github.com/rsp4jack/now_playing">https://github.com/rsp4jack/now_playing</a>
        <hr/>
        """

import asyncio
import ctypes
import ctypes.wintypes
import logging
import site
import sys
import threading
import time
import traceback
from collections import namedtuple
from concurrent.futures import ThreadPoolExecutor
from itertools import chain
from types import LambdaType
from typing import Any, AnyStr, Callable, Sequence

import win32api
import win32con
import win32gui
import win32process
import winrt.windows.foundation as _
from winrt.windows.media.control import \
    GlobalSystemMediaTransportControlsSessionManager as SMTCManager
from winrt.windows.media.control import \
    GlobalSystemMediaTransportControlsSessionMediaProperties as SMTCProperties

import obspython as obs

def timeit(func):
    async def process(func, *args, **params):
        if asyncio.iscoroutinefunction(func):
            return await func(*args, **params)
        else:
            return func(*args, **params)

    async def helper(*args, **params):
        start = time.time()
        result = await process(func, *args, **params)
        end = time.time()

        log.debug(f'{func.__name__} takes {end - start}s')
        return result

    return helper

def IsWindowVisibleOnScreen(hwnd):
    def IsWindowCloaked(hwnd):
        DWMWA_CLOAKED = 14
        cloaked = ctypes.wintypes.DWORD()
        ctypes.windll.dwmapi.DwmGetWindowAttribute(hwnd, ctypes.wintypes.DWORD(
            DWMWA_CLOAKED), ctypes.byref(cloaked), ctypes.sizeof(cloaked))
        return cloaked.value
    return ctypes.windll.user32.IsWindowVisible(hwnd) and (not IsWindowCloaked(hwnd))

enabled = False
check_frequency = 1000  # ms
display_text = ''
source_name = ''
encaptureSet: set[str] = set()

logging.basicConfig(
    format="[{asctime}] [{threadName}/{levelname}]: [{module}]: {message}",
    datefmt="%H:%M:%S",
    style="{",
    level=logging.INFO
)

log = logging.getLogger(__name__)

class Capture(object):
    def __init__(self, id: AnyStr, display_name: AnyStr, func: Callable):
        self._id = id
        self._display_name = display_name
        self._func = func
    
    @property
    def id(self):
        return self._id
    
    @property
    def display_name(self):
        return self._display_name

    def __call__(self, *args, **kwargs):
        return self._func(*args, **kwargs)

def win32TitleCaptureWrapper(process: str):
    def wrapper(func: Callable):
        def wrapped() -> list[dict[str, Any]]:
            def enumHandler(hwnd, result: list[dict[str, Any]]):
                _, procpid = win32process.GetWindowThreadProcessId(hwnd)
                try:
                    if not IsWindowVisibleOnScreen(hwnd):
                        return
                    mypyproc = win32api.OpenProcess(
                        win32con.PROCESS_QUERY_INFORMATION, False, procpid)
                    image: str = win32process.GetModuleFileNameEx(mypyproc, 0)
                    title: str = win32gui.GetWindowText(hwnd)
                    
                    if image.split('\\')[-1] == process:
                        result.extend(func(title))
                except Exception:
                    log.warning(f'enumHandler error for {process}', exc_info=True)
                    return
                return
            result: list[dict[str, Any]] = []
            win32gui.EnumWindows(enumHandler, result)
            return result
        return wrapped
    return wrapper

def foobar2000capture(title: str) -> tuple[str, str]:
    artist = ''
    song = ''
    if ("-" not in title) and (title.find('[foobar2000]') != -1):
        song = title[:title.rfind(" [foobar2000]")-1]
    elif "-" in title:
        artist = title[0:title.find("-")-1]
        song = title[title.find("]")+2:title.rfind(" [foobar2000]")-1]
    return (artist, song)

manager: SMTCManager | None = None

@timeit
async def smtcCaptureAsync() -> list[dict[str, Any]]:
    global manager
    if not manager:
        manager = await SMTCManager.request_async()
    session = manager.get_current_session()
    if not session:
        return []
    try:
        properties: SMTCProperties = await session.try_get_media_properties_async()
    except PermissionError as err:
        if err.winerror == -2147024875:
            log.warning('SMTCSession try_get_media_properties_async(): ERROR_NOT_READY', exc_info=True)
    # TODO: more properties
    # TODO: use smtc event handler
    return [{'artist': properties.artist, 'title': properties.title}]

def smtcCapture() -> list[dict[str, Any]]:
    return asyncio.run_coroutine_threadsafe(smtcCaptureAsync(), loop).result(5)

captures: dict[str, Capture] = {
    'smtc': Capture('smtc', 'SMTC', smtcCapture),
    'spotify': Capture('spotify', 'Spotify', win32TitleCaptureWrapper('spotify.exe')(lambda x: [{'artist': x[0:x.find('-')-1], 'title': x[x.find('-')+2:]}] if '-' in x else [])),
    'vlc': Capture('vlc', "VLC", win32TitleCaptureWrapper('vlc.exe')(lambda x: [{'artist': x[0:x.find('-')-1], 'title': x[x.find('-')+2:x.rfind('-')-1]}] if '-' in x else [])),
    'yt_firefox': Capture('yt_firefox', "YouTube for Firefox", win32TitleCaptureWrapper('firefox.exe')(lambda x: [{'artist': x[0:x.find('-')-1], 'title': x[x.find('-')+2:x.rfind('-')-1]}] if '- YouTube' in x else [])),
    'yt_chrome': Capture('yt_chrome', 'YouTube for Chrome', win32TitleCaptureWrapper('chrome.exe')(lambda x: [{'artist': x[0:x.find('-')-1], 'title': x[x.find('-')+2:x.rfind('-')-1]}] if '- YouTube' in x else [])),
    'foobar2000': Capture('foobar2000', 'foobar2000', win32TitleCaptureWrapper('foobar2000.exe')(foobar2000capture)),
    'necloud': Capture('necloud', 'necloud', win32TitleCaptureWrapper('cloudmusic.exe')(lambda x: [{'artist': x[x.find("-")+2:], 'title': x[0:x.find("-")-1]}] if '-' in x else [])),
    'aimp': Capture('aimp', 'AIMP', win32TitleCaptureWrapper('aimp.exe')(lambda x: [{'artist': x[0:x.find('-')-1], 'title': x[x.find('-')+2:]}] if '-' in x else [])),
}


def script_properties():
    log.debug("script_properties()")
    # log.info(f'locale: {obs.obs_get_locale()}')

    props = obs.obs_properties_create()
    obs.obs_properties_add_bool(props, "enabled", "Enabled")
    logcombo = obs.obs_properties_add_list(props, "log_level", "Log level", obs.OBS_COMBO_TYPE_EDITABLE, obs.OBS_COMBO_FORMAT_STRING)
    for name in ['NOTSET', 'DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITIAL', 'SILENT']:
        obs.obs_property_list_add_string(logcombo, name, name)
    obs.obs_properties_add_int(
        props, "check_frequency", "Check frequency", 150, 60000, 100)
    obs.obs_properties_add_text(
        props, "display_text", "Display text", obs.OBS_TEXT_DEFAULT)
    for name, cap in captures.items():
        obs.obs_properties_add_bool(props, name, cap.display_name)

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

def script_defaults(settings):
    log.debug(f"script_defaults({settings!r})")

    obs.obs_data_set_default_bool(settings, "enabled", True)
    obs.obs_data_set_default_int(settings, "check_frequency", 1000)
    obs.obs_data_set_default_string(settings, "display_text", "%artist - %title")
    obs.obs_data_set_default_string(settings, "source_name", '')
    obs.obs_data_set_default_string(settings, "log_level", 'INFO')
    for name in captures.keys():
        obs.obs_data_set_default_bool(settings, name, name == 'smtc')

def script_save(settings):
    log.debug(f"script_save({settings!r})")

    script_update(settings)

def script_update(settings):
    global enabled
    global display_text
    global check_frequency
    global source_name
    log.debug(f"script_update({settings!r})")

    loglevel = obs.obs_data_get_string(settings, "log_level")
    if loglevel == 'SILENT':
        logging.getLogger().setLevel(logging.CRITICAL + 100)
    else:
        try:
            logging.getLogger().setLevel(loglevel)
        except ValueError:
            traceback.print_exc(file=sys.stderr)
            logging.getLogger().setLevel(logging.INFO)


    display_text = obs.obs_data_get_string(settings, "display_text")
    source_name = obs.obs_data_get_string(settings, "source_name")
    new_check_frequency = obs.obs_data_get_int(settings, "check_frequency")
    encaptureSet.clear()
    for name in captures.keys():
        if obs.obs_data_get_bool(settings, name):
            encaptureSet.add(name)
    
    if new_check_frequency != check_frequency and enabled:
        check_frequency = new_check_frequency
        obs.timer_remove(onUpdate)
        obs.timer_add(onUpdate, check_frequency)
    
    if obs.obs_data_get_bool(settings, "enabled"):
        if not enabled:
            log.debug('Timer started')
            enabled = True
            obs.timer_add(onUpdate, check_frequency)
    else:
        if enabled:
            log.debug('Timer killed')
            enabled = False
            obs.timer_remove(onUpdate)

def update_song(data: dict[str, Any]):
    now_playing = display_text
    for key, value in data.items():
        now_playing = now_playing.replace(f"%{key}", value)

    settings = obs.obs_data_create()
    obs.obs_data_set_string(settings, "text", now_playing)
    source = obs.obs_get_source_by_name(source_name)
    obs.obs_source_update(source, settings)
    obs.obs_data_release(settings)
    obs.obs_source_release(source)

    log.debug(f"updated: {now_playing} <- {data}")

tpool = ThreadPoolExecutor(4, 'nowplaying_updateworker')
loop = asyncio.new_event_loop()
loop.set_default_executor(tpool)
loopthread: threading.Thread | None = None
def start_eventloop():
    asyncio.set_event_loop(loop)
    loop.run_forever()
def startevthread():
    global loopthread
    log.debug('Starting event loop thread')
    if loopthread and loopthread.is_alive():
        log.warning('loopthread is still alive!!!', stack_info=True)
        loop.stop()
    loopthread = threading.Thread(target=start_eventloop, name='nowplaying_eventloop', daemon=True)
    loopthread.start()

def script_load(_):
    log.debug('script_load()')
    startevthread()

def script_unload():
    log.debug('script_unload()')
    obs.timer_remove(onUpdate)
    [task.cancel('plugin unloaded') for task in asyncio.all_tasks(loop)]
    loop.stop()


def onUpdate():
    fut = asyncio.run_coroutine_threadsafe(doUpdate(), loop)
    def callback(f):
        if exc := f.exception():
            log.error('doUpdate fut error', exc_info=exc)
            return
    fut.add_done_callback(callback)

async def doUpdate():
    try:
        datalist = await asyncio.gather(*[asyncio.to_thread(captures[name]) for name in encaptureSet ])
        data: list[dict[str, Any]] = list(chain(*datalist))
    except Exception:
        log.warning('capture error', exc_info=True)
        return
    
    log.debug(f"doUpdate: {data}")
    if not data:
        return
    try:
        update_song(data[0])
    except Exception:
        log.error('update_song error', exc_info=True)
