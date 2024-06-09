def script_description():
    log.debug("script_description()")

    return """
        <h1>smcinfo.py</h1>
        <hr/>
        <a href="https://github.com/rsp4jack/now_playing">https://github.com/rsp4jack/now_playing</a>
        <hr/>
        """


DEFAULT_DISPLAY_EXPR = r"""
'NO MEDIA' if not data else \
''.join([
    f'{artist} - {title} ',
    *(
        ['\n', f'{fmttd(roundtd(predictedpos()))}/{fmttd(roundtd(end_time))}']
        if posavail() else
        []
    )
])
""".strip()

import asyncio
import aiohttp
import concurrent.futures
import logging
import os
import shutil
import sys
import tempfile
import threading
import time
import traceback
from collections.abc import Coroutine
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta, timezone
from types import CodeType
from typing import Any, cast
import platform
import urllib.parse

MEDIACTRL = {'Windows': 'SMTC', 'Linux': 'MPRIS',}.get(platform.system())

if MEDIACTRL == 'SMTC':
    import winrt.windows.foundation as _
    from winrt.windows.foundation import EventRegistrationToken
    from winrt.windows.foundation.collections import IVectorView as _
    from winrt.windows.media.control import CurrentSessionChangedEventArgs
    from winrt.windows.media.control import \
        GlobalSystemMediaTransportControlsSession as SMTCSession
    from winrt.windows.media.control import \
        GlobalSystemMediaTransportControlsSessionManager as SMTCManager
    from winrt.windows.media.control import \
        GlobalSystemMediaTransportControlsSessionMediaProperties as SMTCProperties
    from winrt.windows.media.control import \
        GlobalSystemMediaTransportControlsSessionTimelineProperties as \
        TimelineProperties
    from winrt.windows.media.control import GlobalSystemMediaTransportControlsSessionPlaybackInfo as PlaybackInfo
    from winrt.windows.media.control import (MediaPropertiesChangedEventArgs,
                                            TimelinePropertiesChangedEventArgs, PlaybackInfoChangedEventArgs)
    from winrt.windows.storage.streams import (DataReader,
                                            IRandomAccessStreamReference)
elif MEDIACTRL == 'MPRIS':
    from dbus_next.aio.message_bus import MessageBus, ProxyObject
    from dbus_next.message import Message
    from dbus_next.constants import MessageType

import obspython as obs

def convert_future_exc(exc):
    exc_class = type(exc)
    if exc_class is concurrent.futures.CancelledError:
        return asyncio.CancelledError(*exc.args).with_traceback(exc.__traceback__)
    elif exc_class is concurrent.futures.InvalidStateError:
        return asyncio.InvalidStateError(*exc.args).with_traceback(exc.__traceback__)
    else:
        return exc

asyncio.futures._convert_future_exc = convert_future_exc  # type: ignore

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

        log.debug(f"{func.__name__} takes {end - start}s")
        return result

    return helper

logging.basicConfig(
    format="[{asctime}] [{module}] [{threadName}/{levelname}]: [{funcName}]: {message}",
    datefmt="%H:%M:%S",
    style="{",
    level=logging.INFO,
)

log = logging.getLogger(__name__)

log.info(f'Using media control: {MEDIACTRL}')

###! <---
###! PROP
###! --->

enabled = False
update_frequency = 1000 # ms
display_expr: CodeType | None = None
source_name = ""
thumbsource_name = ""

media_props: dict[str, Any] = {}
timeline_props: dict[str, Any] = {}

def script_properties():
    log.debug("script_properties()")
    # log.info(f'locale: {obs.obs_get_locale()}')

    props = obs.obs_properties_create()
    obs.obs_properties_add_bool(props, "enabled", "Enabled")
    logcombo = obs.obs_properties_add_list(
        props,
        "log_level",
        "Log level",
        obs.OBS_COMBO_TYPE_EDITABLE,
        obs.OBS_COMBO_FORMAT_STRING,
    )
    for name in ["NOTSET", "DEBUG", "INFO", "WARNING", "ERROR", "CRITIAL", "SILENT"]:
        obs.obs_property_list_add_string(logcombo, name, name)
    obs.obs_properties_add_text(
        props, "display_expr", "Display expr", obs.OBS_TEXT_MULTILINE
    )

    p = obs.obs_properties_add_list(
        props,
        "source_name",
        "Text source",
        obs.OBS_COMBO_TYPE_EDITABLE,
        obs.OBS_COMBO_FORMAT_STRING,
    )
    p2 = obs.obs_properties_add_list(
        props,
        "thumbsource_name",
        "Thumbnail source",
        obs.OBS_COMBO_TYPE_EDITABLE,
        obs.OBS_COMBO_FORMAT_STRING,
    )

    sources = obs.obs_enum_sources()
    if sources:
        for source in sources:
            source_id = obs.obs_source_get_unversioned_id(source)
            if source_id in ("text_gdiplus", "text_ft2_source"):
                name = obs.obs_source_get_name(source)
                obs.obs_property_list_add_string(p, name, name)
            elif source_id == "image_source":
                name = obs.obs_source_get_name(source)
                obs.obs_property_list_add_string(p2, name, name)
    obs.source_list_release(sources)

    return props


def script_defaults(settings):
    log.debug(f"script_defaults({settings!r})")

    obs.obs_data_set_default_bool(settings, "enabled", True)
    obs.obs_data_set_default_string(settings, "display_expr", DEFAULT_DISPLAY_EXPR)
    obs.obs_data_set_default_string(settings, "source_name", "")
    obs.obs_data_set_default_string(settings, "thumbsource_name", "")
    obs.obs_data_set_default_string(settings, "log_level", "INFO")


def script_save(settings):
    log.debug(f"script_save({settings!r})")

    script_update(settings)



def script_update(settings):
    global enabled
    global display_expr
    global check_frequency
    global source_name
    global thumbsource_name
    log.debug(f"script_update({settings!r})")

    loglevel = obs.obs_data_get_string(settings, "log_level")
    if loglevel == "SILENT":
        logging.getLogger().setLevel(logging.CRITICAL + 100)
    else:
        try:
            logging.getLogger().setLevel(loglevel)
        except ValueError:
            traceback.print_exc(file=sys.stderr)
            logging.getLogger().setLevel(logging.INFO)

    display_expr = compile(
        obs.obs_data_get_string(settings, "display_expr"), "<string>", "eval"
    )
    source_name = obs.obs_data_get_string(settings, "source_name")
    thumbsource_name = obs.obs_data_get_string(settings, "thumbsource_name")

    toenabled = obs.obs_data_get_bool(settings, "enabled")
    if toenabled and not enabled:
        log.info('Initalizing media controls')
        runcoro(smcInitalizeAsync())
        obs.timer_add(on_timer, 500)
    elif not toenabled and enabled:
        log.info('Deinitalizing media controls')
        runcoro(smcDeinitalizeAsync())
        obs.timer_remove(on_timer)
    enabled = toenabled
    if enabled:
        runcoro(smcUpdateAsync())

###! <---
###! SMTC
###! --->

lastData: dict[str, Any] | None = None

if MEDIACTRL == 'SMTC':
    manager: SMTCManager | None = None
    onSessionChangedToken: EventRegistrationToken | None = None
    currentSession: SMTCSession | None = None
    onMediaPropChangedToken: EventRegistrationToken | None = None
    onTimelineChangedToken: EventRegistrationToken | None = None
    onPlaybackInfoChangedToken: EventRegistrationToken | None = None
    thumbdir: str | None = None

    async def smtcDeinitalizeAsync():
        global manager
        global onSessionChangedToken
        await smtcSetSessionAsync(None)
        if manager and onSessionChangedToken:
            manager.remove_current_session_changed(onSessionChangedToken)
            onSessionChangedToken = None
        manager = None


    async def smtcInitalizeAsync():
        global manager
        global onSessionChangedToken
        await smtcDeinitalizeAsync()
        manager = await SMTCManager.request_async()

        def onSessionChanged(
            sender: SMTCManager | None, event: CurrentSessionChangedEventArgs | None
        ):
            assert sender
            session = sender.get_current_session()
            runcoro(smtcSetSessionAsync(session))

        onSessionChangedToken = manager.add_current_session_changed(onSessionChanged)
        await smtcSetSessionAsync(manager.get_current_session())


    async def smtcSetSessionAsync(session: SMTCSession | None):
        global currentSession
        global onMediaPropChangedToken
        global onTimelineChangedToken
        global onPlaybackInfoChangedToken
        log.debug(f"smtcSetSession(): {currentSession!r} -> {session!r}")
        if currentSession:
            if onMediaPropChangedToken:
                currentSession.remove_media_properties_changed(onMediaPropChangedToken)
                onMediaPropChangedToken = None
            if onTimelineChangedToken:
                currentSession.remove_timeline_properties_changed(onTimelineChangedToken)
                onTimelineChangedToken = None
            if onPlaybackInfoChangedToken:
                currentSession.remove_playback_info_changed(onPlaybackInfoChangedToken)
                onPlaybackInfoChangedToken = None
        currentSession = session
        if not currentSession:
            return

        def onMediaPropChanged(
            sender: SMTCSession | None, event: MediaPropertiesChangedEventArgs | None
        ):
            runcoro(smtcUpdateAsync(currentSession))

        onMediaPropChangedToken = currentSession.add_media_properties_changed(
            onMediaPropChanged
        )

        def onTimelineChanged(
            sender: SMTCSession | None, event: TimelinePropertiesChangedEventArgs | None
        ):
            runcoro(smtcUpdateAsync(currentSession, thumb=False))

        onTimelineChangedToken = currentSession.add_timeline_properties_changed(
            onTimelineChanged
        )

        def onPlaybackInfoChanged(
            sender: SMTCSession | None, event: PlaybackInfoChangedEventArgs | None
        ):
            runcoro(smtcUpdateAsync(currentSession, thumb=False))
        
        onPlaybackInfoChangedToken = currentSession.add_playback_info_changed(onPlaybackInfoChanged)

        await smtcUpdateAsync(currentSession)


    async def smtcUpdateAsync(session: SMTCSession | None, *, thumb: bool = True, capture: bool = True):
        global lastData
        if capture:
            datas = await smtcCaptureAsync(session)
            data = datas[0] if datas else None
            lastData = data
        else:
            data = lastData
        update_text(data)
        if not data or not data.get('thumbnail'):
            update_thumbnail('')
        elif thumb:
            file = await fetch_thumbnail_async(data.get("thumbnail"))
            update_thumbnail(file)

    @timeit
    async def smtcCaptureAsync(session: SMTCSession | None) -> list[dict[str, Any]]:
        if not session:
            return []
        try:
            properties: SMTCProperties = await session.try_get_media_properties_async()
            timeline: TimelineProperties | None = session.get_timeline_properties()
            playback: PlaybackInfo | None = session.get_playback_info()
        except PermissionError as err:
            if err.winerror == -2147024875:
                log.info(
                    "SMTCSession try_get_media_properties_async(): ERROR_NOT_READY",
                    exc_info=True,
                )
                return []
        # TODO: more properties
        # TODO: use smtc event handler
        mediaprop = {
            "artist": properties.artist,
            "title": properties.title,
            "subtitle": properties.subtitle,
            "track_number": properties.track_number,
            "genres": list(properties.genres) if properties.genres else None,
            "album_title": properties.album_title,
            "album_artist": properties.album_artist,
            "album_track_count": properties.album_track_count,
            "thumbnail": properties.thumbnail,
        }
        timelineprop = {}
        if timeline:
            timelineprop = {
                "position": timeline.position,
                "last_updated_time": timeline.last_updated_time,
                "start_time": timeline.start_time,
                "end_time": timeline.end_time,
                "min_seek_time": timeline.min_seek_time,
                "max_seek_time": timeline.max_seek_time,
            }
        playbackprop = {}
        if playback:
            playbackprop = {
                'playback_type': playback.playback_type.name.capitalize() if playback.playback_type else None,
                'playback_rate': playback.playback_rate,
                'playback_status': playback.playback_status.name.capitalize(),
                'repeat_mode': playback.auto_repeat_mode.name.capitalize() if playback.auto_repeat_mode else None,
                'is_shuffle_active': playback.is_shuffle_active
            }
        log.debug(f"captured: {mediaprop}, {timelineprop} {playbackprop}")
        return [{**mediaprop, **timelineprop, **playbackprop}]
    
    async def fetch_thumbnail_async(thumb: IRandomAccessStreamReference) -> str:
        assert thumbdir
        with await thumb.open_read_async() as rastream:
            log.debug(f"received thumb {rastream.content_type} {rastream.size}bytes")
            with open(os.path.join(thumbdir, "thumbnail"), "wb") as f:
                log.debug(f"update_thumbnai_async mkstemp {f.name}")
                with DataReader(rastream.get_input_stream_at(0)) as reader:
                    await reader.load_async(rastream.size)
                    while True:
                        len = min(8192, reader.unconsumed_buffer_length)
                        if len == 0:
                            break
                        try:
                            buf = reader.read_buffer(len)
                        except:
                            log.warning(f"read_buffer({len}) failed", exc_info=True)
                            break
                        if not buf:
                            break
                        written = f.write(buf)
                        log.debug(
                            f"thumb written {written}, buf {buf.length}, read {len}"
                        )
                return f.name
    
    smcInitalizeAsync = smtcInitalizeAsync
    smcDeinitalizeAsync = smtcDeinitalizeAsync
    async def smcUpdateAsync(*args, **kwargs):
        return await smtcUpdateAsync(currentSession, *args, **kwargs)

###! <---
###! MPRIS
###! --->

elif MEDIACTRL == 'MPRIS':
    bus: MessageBus | None = None
    playerobj: ProxyObject | None = None
    
    async def mprisInitalize():
        global bus
        log.info('Initalizing DBus')
        bus = await MessageBus().connect()
        await mprisDiscoverService()
    
    async def mprisDiscoverService():
        global playerobj
        assert(bus)

        reply = await bus.call(
        Message(destination='org.freedesktop.DBus',
                path='/org/freedesktop/DBus',
                interface='org.freedesktop.DBus',
                member='ListNames'))
        assert(reply)
        if reply.message_type == MessageType.ERROR:
            raise RuntimeError(reply.body[0])

        services: list[str] = reply.body[0]
        players = [s for s in services if s.startswith('org.mpris.MediaPlayer2.')]
        if not players:
            log.debug('No MPRIS players found')
            playerobj = None
            return
        if 'org.mpris.MediaPlayer2.playerctld' in players:
            busname = 'org.mpris.MediaPlayer2.playerctld'
        else:
            busname = players[0]
        log.info(f'Using MPRIS bus {busname}')
        
        introspect = await bus.introspect(busname, '/org/mpris/MediaPlayer2')
        playerobj = bus.get_proxy_object(busname, '/org/mpris/MediaPlayer2', introspect)

        playeriface = playerobj.get_interface('org.mpris.MediaPlayer2.Player')
        propiface = playerobj.get_interface('org.freedesktop.DBus.Properties')

        async def on_properties_changed(interface, changed, invalidated):
            log.debug(f'MPRIS on_properties_changed: {interface!r} {changed!r} {invalidated!r}')
            await mprisUpdate()
        
        async def on_seeked(pos):
            log.debug(f'MPRIS on_seeked: {pos}')
            await mprisUpdate()

        propiface.on_properties_changed(on_properties_changed) # type: ignore
        playeriface.on_seeked(on_seeked) # type: ignore

    @timeit
    async def mprisCapture():
        assert(playerobj)
        player = playerobj.get_interface('org.mpris.MediaPlayer2.Player')
        meta: dict[str, Any] = await player.get_metadata() # type: ignore
        meta = {k.lower(): v.value for k,v in meta.items()}
        position = timedelta(microseconds=await player.get_position()) # type: ignore
        data = {
            'artist': ', '.join(meta.get('xesam:artist', [])),
            'title': meta.get('xesam:title'),
            'track_number': meta.get('xesam:tracknumber'),
            'genres': meta.get('xesam:genre'),
            'album_title': meta.get('xesam:album'),
            'album_artist': meta.get('xesam:albumartist'),
            'album_track_count': meta.get('xesam:albumtrackcount'),
            'thumbnail': await mprisFetchThumbnail(meta['mpris:arturl']) if 'mpris:arturl' in meta else None,

            'position': position,
            'end_time': timedelta(microseconds=meta['mpris:length']),
            'last_updated_time': datetime.now(timezone.utc),

            'playback_status': await player.get_playback_status(), # type: ignore
            'repeat_mode': await player.get_loop_status(), # type: ignore
            'playback_rate': await player.get_rate(), # type: ignore
        }
        log.debug(f"captured: {data}")
        return [data]
    
    async def mprisUpdate(*, thumb: bool = True, capture: bool = True):
        global lastData
        if capture:
            datas = await mprisCapture()
            data = datas[0] if datas else None
            lastData = data
        else:
            data = lastData
        update_text(data)
        if not data or not data.get('thumbnail'):
            update_thumbnail('')
        elif thumb:
            file = data["thumbnail"]
            update_thumbnail(file)
    
    async def mprisFetchThumbnail(url: str):
        parsed = urllib.parse.urlparse(url)
        if parsed.scheme == 'file':
            return urllib.parse.unquote_plus(parsed.path)
        if parsed.scheme == 'http' or parsed.scheme == 'https':
            assert(thumbdir)
            async with aiohttp.ClientSession(trust_env=True) as session:
                async with session.get(url) as resp:
                    resp.raise_for_status()
                    with open(os.path.join(thumbdir, url.replace('/', '_')), 'wb') as f:
                        f.write(await resp.read())
                        return f.name
        raise ValueError(f"Unsupported thumbnail URL: {url}")
    
    smcInitalizeAsync = mprisInitalize
    async def smcDeinitalizeAsync():
        pass
    smcUpdateAsync = mprisUpdate
    

def update_text(data: dict[str, Any] | None):
    def roundtd(td: timedelta) -> timedelta:
        return timedelta(seconds=round(td.total_seconds()))

    def fmttd(td: timedelta):
        return str(td).removeprefix("0:").removeprefix("0")

    def posavail():
        return data and (
            "last_updated_time" in data
            and cast(datetime, data["last_updated_time"]).year != 1601
        )

    def predictedpos() -> timedelta:
        assert(data)
        if data['playback_status'] != 'Playing':
            return data['position']
        return data['position']+(datetime.now(timezone.utc)-data['last_updated_time'])*data['playback_rate']

    namespace: dict[str, Any] = {
        "data": data,
        "roundtd": roundtd,
        "fmttd": fmttd,
        "posavail": posavail,
        "predictedpos": predictedpos,
    }
    namespace.update(sys.modules)
    if data:
        namespace.update(data)
    if display_expr:
        try:
            now_playing = eval(display_expr, namespace)
        except:
            log.warning("Failed to evaluate display expression", exc_info=True)
            now_playing = "..."
    else:
        log.warning('No display expression')
        now_playing = "..."
    settings = obs.obs_data_create()
    obs.obs_data_set_string(settings, "text", now_playing)
    source = obs.obs_get_source_by_name(source_name)
    obs.obs_source_update(source, settings)
    obs.obs_data_release(settings)
    obs.obs_source_release(source)

    log.debug(f"source {source_name}: {now_playing} <- {data}")

def update_thumbnail(file: str):
    props = obs.obs_data_create()
    obs.obs_data_set_string(props, "file", file)
    thumbsrc = obs.obs_get_source_by_name(thumbsource_name)
    obs.obs_source_update(thumbsrc, props)
    obs.obs_data_release(props)
    obs.obs_source_release(thumbsrc)

    log.debug(
        f"source {thumbsource_name}: {file}"
    )


###! <---
###! SCHED
###! --->


tpool = ThreadPoolExecutor(4, "smc_pool_thread_")
loop = asyncio.new_event_loop()
loop.set_default_executor(tpool)
loopthread: threading.Thread | None = None

def startevthread():
    global loopthread
    log.debug("Starting event loop thread")
    if loopthread and loopthread.is_alive():
        log.warning("loopthread is still alive!!!", stack_info=True)
        loop.stop()
    loopthread = threading.Thread(
        target=loop.run_forever, name="smc_evloop", daemon=True
    )
    asyncio.set_event_loop(loop)
    loopthread.start()


def runcoro(coro: Coroutine, timeout: float | None = None):
    fut = asyncio.run_coroutine_threadsafe(coro, loop)
    return fut.result(timeout)


def script_load(_):
    global thumbdir
    log.debug("script_load()")
    thumbdir = tempfile.mkdtemp(prefix="smcinfo_thumbs_")
    startevthread()

def script_unload():
    global thumbdir
    log.debug("script_unload()")
    if thumbdir:
        shutil.rmtree(thumbdir)
        thumbdir = None
    runcoro(smcDeinitalizeAsync(), 5)
    [task.cancel("plugin unloaded") for task in asyncio.all_tasks(loop)]
    loop.stop()


def on_timer():
    runcoro(smcUpdateAsync(thumb=False, capture=False))
