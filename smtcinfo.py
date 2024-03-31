def script_description():
    log.debug("script_description()")

    return \
        """
        <h1>smtcinfo.py</h1>
        <hr/>
        <a href="https://github.com/rsp4jack/now_playing">https://github.com/rsp4jack/now_playing</a>
        <hr/>
        """

import asyncio
from collections.abc import Coroutine
from datetime import timedelta
import logging
import site
import sys
import threading
import time
import traceback
from collections import namedtuple
from concurrent.futures import ThreadPoolExecutor
from itertools import chain
from types import CodeType, LambdaType
from typing import Any, AnyStr, Callable, Sequence
import concurrent.futures

from winrt.windows.foundation import EventRegistrationToken
import winrt.windows.foundation as _
from winrt.windows.foundation.collections import IVectorView
from winrt.windows.media.control import \
    GlobalSystemMediaTransportControlsSessionManager as SMTCManager
from winrt.windows.media.control import \
    GlobalSystemMediaTransportControlsSession as SMTCSession
from winrt.windows.media.control import GlobalSystemMediaTransportControlsSessionTimelineProperties as TimelineProperties
from winrt.windows.media.control import \
    GlobalSystemMediaTransportControlsSessionMediaProperties as SMTCProperties
from winrt.windows.media.control import CurrentSessionChangedEventArgs, TimelinePropertiesChangedEventArgs, MediaPropertiesChangedEventArgs

import obspython as obs

def convert_future_exc(exc):
    exc_class = type(exc)
    if exc_class is concurrent.futures.CancelledError:
        return asyncio.CancelledError(*exc.args).with_traceback(exc.__traceback__)
    elif exc_class is concurrent.futures.TimeoutError:
        return asyncio.TimeoutError(*exc.args).with_traceback(exc.__traceback__)
    elif exc_class is concurrent.futures.InvalidStateError:
        return asyncio.InvalidStateError(*exc.args).with_traceback(exc.__traceback__)
    else:
        return exc

asyncio.futures._convert_future_exc = convert_future_exc # type: ignore

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


enabled = False
check_frequency = 1000  # ms
display_expr: CodeType | None = None
source_name = ''

logging.basicConfig(
    format="[{asctime}] [{threadName}/{levelname}]: [{module}]: {message}",
    datefmt="%H:%M:%S",
    style="{",
    level=logging.INFO
)

log = logging.getLogger(__name__)


def script_properties():
    log.debug("script_properties()")
    # log.info(f'locale: {obs.obs_get_locale()}')

    props = obs.obs_properties_create()
    obs.obs_properties_add_bool(props, "enabled", "Enabled")
    logcombo = obs.obs_properties_add_list(props, "log_level", "Log level", obs.OBS_COMBO_TYPE_EDITABLE, obs.OBS_COMBO_FORMAT_STRING)
    for name in ['NOTSET', 'DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITIAL', 'SILENT']:
        obs.obs_property_list_add_string(logcombo, name, name)
    obs.obs_properties_add_text(
        props, "display_expr", "Display expr", obs.OBS_TEXT_DEFAULT)

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
    obs.obs_data_set_default_string(settings, "display_expr", "f'{data[\"artist\"]} - {data[\"title\"]}'")
    obs.obs_data_set_default_string(settings, "source_name", '')
    obs.obs_data_set_default_string(settings, "log_level", 'INFO')

def script_save(settings):
    log.debug(f"script_save({settings!r})")

    script_update(settings)

def script_update(settings):
    global enabled
    global display_expr
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


    display_expr = compile(obs.obs_data_get_string(settings, "display_expr"), '<string>', 'eval')
    source_name = obs.obs_data_get_string(settings, "source_name")
    
    enabled = obs.obs_data_get_bool(settings, "enabled")
    if enabled and not manager:
        runcoro(smtcInitalizeAsync())
    elif not enabled and manager:
        runcoro(smtcDeinitalizeAsync())

    if enabled:
        smtcUpdate(currentSession)

manager: SMTCManager | None = None
onSessionChangedToken: EventRegistrationToken | None = None
currentSession: SMTCSession | None = None
onMediaPropChangedToken: EventRegistrationToken | None = None
onTimelineChangedToken: EventRegistrationToken | None = None

async def smtcDeinitalizeAsync():
    global manager
    global onSessionChangedToken
    smtcSetSession(None)
    if manager and onSessionChangedToken:
        manager.remove_current_session_changed(onSessionChangedToken)
        onSessionChangedToken = None
    manager = None

async def smtcInitalizeAsync():
    global manager
    global onSessionChangedToken
    await smtcDeinitalizeAsync()
    manager = await SMTCManager.request_async()

    def onSessionChanged(sender: SMTCManager | None, event: CurrentSessionChangedEventArgs | None):
        assert(sender)
        session = sender.get_current_session()
        smtcSetSession(session)

    onSessionChangedToken = manager.add_current_session_changed(onSessionChanged)
    tpool.submit(onSessionChanged, manager, None)

def smtcSetSession(session: SMTCSession | None):
    global currentSession
    global onMediaPropChangedToken
    global onTimelineChangedToken
    log.debug(f'smtcSetSession(): {currentSession!r} -> {session!r}')
    if currentSession:
        if onMediaPropChangedToken:
            currentSession.remove_media_properties_changed(onMediaPropChangedToken)
            onMediaPropChangedToken = None
        if onTimelineChangedToken:
            currentSession.remove_timeline_properties_changed(onTimelineChangedToken)
            onTimelineChangedToken = None
    currentSession = session
    if not currentSession:
        return
    
    def onMediaPropChanged(sender: SMTCSession | None, event: MediaPropertiesChangedEventArgs | None):
        smtcUpdate(currentSession)
    onMediaPropChangedToken = currentSession.add_media_properties_changed(onMediaPropChanged)

    def onTimelineChanged(sender: SMTCSession | None, event: TimelinePropertiesChangedEventArgs | None):
        smtcUpdate(currentSession)
    onTimelineChangedToken = currentSession.add_timeline_properties_changed(onTimelineChanged)

    smtcUpdate(currentSession)

def smtcUpdate(session: SMTCSession | None):
    datas = smtcCapture(session)
    if not datas:
        log.debug('smtcUpdate(): no session')
        return

    update_text(datas[0])

def smtcCapture(session: SMTCSession | None, timeout: float = 3) -> list[dict[str, Any]]:
    return runcoro(smtcCaptureAsync(session), timeout)

@timeit
async def smtcCaptureAsync(session: SMTCSession | None) -> list[dict[str, Any]]:
    if not session:
        return []
    try:
        properties: SMTCProperties = await session.try_get_media_properties_async()
        timeline: TimelineProperties | None = session.get_timeline_properties()
    except PermissionError as err:
        if err.winerror == -2147024875:
            log.warning('SMTCSession try_get_media_properties_async(): ERROR_NOT_READY', exc_info=True)
            return []
    # TODO: more properties
    # TODO: use smtc event handler
    mediaprop = {
        'artist': properties.artist,
        'title': properties.title, 
        'subtitle': properties.subtitle,
        'track_number': properties.track_number,
        'genres': list(properties.genres) if properties.genres else None,
        'album_title': properties.album_title,
        'album_artist': properties.album_artist,
        'album_track_count': properties.album_track_count
    }
    timelineprop = {}
    if timeline:
        timelineprop = {
            'position': timeline.position,
            'last_updated_time': timeline.last_updated_time,
            'start_time': timeline.start_time,
            'end_time': timeline.end_time,
            'min_seek_time': timeline.min_seek_time,
            'max_seek_time': timeline.max_seek_time
        }
    log.debug(f'captured: {mediaprop}, {timelineprop}')
    return [{**mediaprop, **timelineprop}]


def update_text(data: dict[str, Any]):
    def roundtimedelta(td: timedelta) -> timedelta:
        return timedelta(seconds=round(td.total_seconds()))
    namespace: dict[str, Any] = {'data': data, 'roundtimedelta': roundtimedelta}
    namespace.update(sys.modules)
    namespace.update(data)
    if display_expr:
        now_playing = eval(display_expr, namespace)
    else:
        now_playing = "..."
    settings = obs.obs_data_create()
    obs.obs_data_set_string(settings, "text", now_playing)
    source = obs.obs_get_source_by_name(source_name)
    obs.obs_source_update(source, settings)
    obs.obs_data_release(settings)
    obs.obs_source_release(source)

    log.debug(f"updated: {now_playing} <- {data}")

tpool = ThreadPoolExecutor(4, 'nowplaying_updateworker')
loop = asyncio.new_event_loop()
loopthread: threading.Thread | None = None

loop.set_default_executor(tpool)

def run_eventloop():
    asyncio.set_event_loop(loop)
    loop.run_forever()

def startevthread():
    global loopthread
    log.debug('Starting event loop thread')
    if loopthread and loopthread.is_alive():
        log.warning('loopthread is still alive!!!', stack_info=True)
        loop.stop()
    loopthread = threading.Thread(target=run_eventloop, name='nowplaying_eventloop', daemon=True)
    loopthread.start()

def runcoro(coro: Coroutine, timeout: float | None = None):
    fut = asyncio.run_coroutine_threadsafe(coro, loop)
    return fut.result(timeout)

def script_load(_):
    log.debug('script_load()')
    startevthread()
    runcoro(smtcInitalizeAsync())
def script_unload():
    log.debug('script_unload()')
    runcoro(smtcDeinitalizeAsync(), 5)
    [task.cancel('plugin unloaded') for task in asyncio.all_tasks(loop)]
    loop.stop()
    

