"""
Local Windows Music Player MCP Server

IMPORTANT:
All tools in this server control the LOCAL WINDOWS COMPUTER where this
MCP server is running.

They do NOT control the XiaoZhi device itself.

This server uses the classic Windows Media Player COM interface for playback.
"""

import os
import random
import threading
import time
import queue
import logging
from pathlib import Path

import pythoncom
import win32com.client

# MCP Python SDK 2.x:
# FastMCP was renamed to MCPServer and moved to mcp.server.
from mcp.server import MCPServer


# ============================================================
# Configuration
# ============================================================

DEFAULT_MUSIC_DIRECTORY = r"<YOUR_MUSIC_DIRECTORY>"

SUPPORTED_EXTENSIONS = {
    ".mp3",
    ".flac",
    ".wav",
    ".m4a",
    ".aac",
    ".ogg",
    ".wma",
    ".opus",
}


# ============================================================
# Logging
# ============================================================

# MCP stdio uses stdout for protocol communication.
# Therefore all logs must go to stderr.

logging.basicConfig(
    level=logging.INFO,
    format="[MusicPlayer] %(asctime)s - %(levelname)s - %(message)s",
)

logger = logging.getLogger("music_player")


# ============================================================
# MCP
# ============================================================

# MCP Python SDK 2.x equivalent of FastMCP("Local Music Player").
mcp = MCPServer("Local Music Player")


# ============================================================
# Global Playlist State
# ============================================================

music_playlist = []
current_index = -1

playlist_directory = DEFAULT_MUSIC_DIRECTORY

# Normal playlist automatic playback
playlist_auto_advance = False

# Random playback state
random_play_enabled = False
random_play_running = False

# Number of songs requested in current random playback
random_play_count = 0

# Number of songs already played in current random playback
random_play_played = 0


# ============================================================
# Helper Functions
# ============================================================

def _get_relative_music_path(song_path: Path) -> str:
    """
    Return a path relative to DEFAULT_MUSIC_DIRECTORY.

    Example:
        E:\\音乐\\周杰伦\\晴天.mp3

    becomes:
        周杰伦\\晴天.mp3
    """

    try:
        root_path = Path(DEFAULT_MUSIC_DIRECTORY).resolve()
        resolved_song = song_path.resolve()

        return str(resolved_song.relative_to(root_path))

    except ValueError:
        return song_path.name


def _resolve_music_path(song_path: str) -> Path:
    """
    Resolve a user-provided music path.

    Both absolute paths and paths relative to DEFAULT_MUSIC_DIRECTORY
    are supported.
    """

    path = Path(song_path).expanduser()

    if not path.is_absolute():
        path = Path(DEFAULT_MUSIC_DIRECTORY) / path

    return path.resolve()


def _song_info(song_path: Path) -> dict:
    """
    Return user-friendly information about a song.
    """

    return {
        "name": song_path.name,
        "relative_path": _get_relative_music_path(song_path),
        "full_path": str(song_path.resolve()),
    }


# ============================================================
# Music Scanner
# ============================================================

def scan_music_directory(directory: str = DEFAULT_MUSIC_DIRECTORY):
    """
    Recursively scan a music directory.

    Songs in subdirectories are also included.
    """

    music_dir = Path(directory).expanduser().resolve()

    if not music_dir.exists():
        raise FileNotFoundError(
            f"Music directory does not exist: {music_dir}"
        )

    if not music_dir.is_dir():
        raise NotADirectoryError(
            f"Not a directory: {music_dir}"
        )

    songs = []

    for path in music_dir.rglob("*"):
        if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS:
            songs.append(path.resolve())

    songs.sort(key=lambda p: str(p).lower())

    return songs


# ============================================================
# Windows Media Player Controller
# ============================================================

class WMPController:
    """
    Dedicated Windows Media Player COM worker.

    The WMP COM object is created and controlled from one dedicated
    COM thread.

    Playback-end detection uses several mechanisms:

    1. WMP MediaEnded state (8)
    2. Previous playback position reaching the duration
    3. Accumulated actual playing time

    The third method is important because some WMP/codec combinations
    jump directly from Playing(3) to Stopped(1) while resetting the
    current position to 0.
    """

    def __init__(self):
        self.command_queue = queue.Queue()

        self.thread = threading.Thread(
            target=self._worker,
            daemon=True,
            name="WMPController",
        )

        self.ready_event = threading.Event()
        self.running = True

        self.wmp = None

        # --------------------------------------------------------
        # Playback state
        # --------------------------------------------------------

        self._media_end_handled = False
        self._seen_playing = False

        self._last_position = 0.0
        self._last_duration = 0.0

        self._last_logged_state = None

        # --------------------------------------------------------
        # Reliable elapsed-playback tracking
        # --------------------------------------------------------

        self._playing_time = 0.0
        self._last_check_time = None
        self._last_state = None

        self.thread.start()

        # Wait for COM/WMP initialization
        self.ready_event.wait(timeout=10)

    # ============================================================
    # Worker
    # ============================================================

    def _worker(self):
        """
        Worker thread owning the WMP COM object.
        """

        pythoncom.CoInitialize()

        try:
            logger.info(
                "Initializing Windows Media Player COM..."
            )

            try:
                self.wmp = win32com.client.Dispatch(
                    "WMPlayer.OCX.7"
                )
            except Exception:
                logger.warning(
                    "WMPlayer.OCX.7 failed, "
                    "trying WMPlayer.OCX..."
                )

                self.wmp = win32com.client.Dispatch(
                    "WMPlayer.OCX"
                )

            # Automatically start playback when media is opened.
            try:
                self.wmp.settings.autoStart = True
            except Exception:
                pass

            # Disable WMP error dialogs.
            try:
                self.wmp.settings.enableErrorDialogs = False
            except Exception:
                pass

            logger.info(
                "Windows Media Player COM initialized."
            )

            self.ready_event.set()

            while self.running:

                # ------------------------------------------------
                # Process queued commands
                # ------------------------------------------------

                try:
                    while True:
                        command, args = (
                            self.command_queue.get_nowait()
                        )

                        try:
                            command(*args)

                        except Exception as e:
                            logger.exception(
                                "WMP command failed: %s",
                                e,
                            )

                except queue.Empty:
                    pass

                # ------------------------------------------------
                # Check playback state
                # ------------------------------------------------

                try:
                    self._check_playback_state()

                except Exception as e:
                    logger.exception(
                        "Playback state check failed: %s",
                        e,
                    )

                # ------------------------------------------------
                # Process COM messages
                # ------------------------------------------------

                try:
                    pythoncom.PumpWaitingMessages()

                except Exception:
                    pass

                time.sleep(0.2)

        finally:

            self.wmp = None

            pythoncom.CoUninitialize()

            logger.info(
                "Windows Media Player COM worker stopped."
            )

    # ============================================================
    # Playback State / EOF Detection
    # ============================================================

    def _check_playback_state(self):
        """
        Detect whether the current song has finished.

        WMP is inconsistent across codecs.

        For example, a song may transition:

            Playing(3)
                ↓
            Stopped(1)
            position = 0

        instead of reporting:

            MediaEnded(8)

        Therefore we use WMP state, position, and accumulated
        actual playing time together.
        """

        if self.wmp is None:
            return

        # --------------------------------------------------------
        # Read WMP state
        # --------------------------------------------------------

        try:
            state = int(self.wmp.playState)

        except Exception:
            return

        # --------------------------------------------------------
        # Read current position
        # --------------------------------------------------------

        position = None

        try:
            position = float(
                self.wmp.controls.currentPosition
            )

        except Exception:
            pass

        # --------------------------------------------------------
        # Read duration
        # --------------------------------------------------------

        duration = None

        try:
            duration = float(
                self.wmp.currentMedia.duration
            )

        except Exception:
            pass

        if duration is not None and duration > 0:
            self._last_duration = duration

        if position is not None and position >= 0:
            current_position = position
        else:
            current_position = self._last_position

        # --------------------------------------------------------
        # Calculate actual playing time
        #
        # Only accumulate time while WMP reports Playing(3).
        # Therefore a manual pause does NOT consume playback time.
        # --------------------------------------------------------

        now = time.monotonic()

        if self._last_check_time is None:
            self._last_check_time = now

        delta = now - self._last_check_time

        if delta < 0:
            delta = 0

        if self._last_state == 3:
            self._playing_time += delta

        self._last_check_time = now

        # --------------------------------------------------------
        # Log state changes
        # --------------------------------------------------------

        if state != self._last_logged_state:

            logger.info(
                "WMP state=%s, position=%.2f, duration=%.2f, "
                "played_time=%.2f",
                state,
                current_position,
                self._last_duration,
                self._playing_time,
            )

            self._last_logged_state = state

        # --------------------------------------------------------
        # Save current state
        # --------------------------------------------------------

        previous_position = self._last_position
        previous_duration = self._last_duration
        previous_seen_playing = self._seen_playing

        self._last_position = current_position
        self._last_state = state

        # --------------------------------------------------------
        # Playing
        # --------------------------------------------------------

        if state == 3:

            self._seen_playing = True

            self._media_end_handled = False

            return

        # --------------------------------------------------------
        # Already handled
        # --------------------------------------------------------

        if self._media_end_handled:
            return

        # --------------------------------------------------------
        # Method 1:
        # Explicit WMP MediaEnded
        # --------------------------------------------------------

        if state == 8:

            logger.info(
                "WMP reported MediaEnded."
            )

            self._media_end_handled = True
            self._seen_playing = False

            self._handle_media_ended()

            return

        # --------------------------------------------------------
        # Method 2:
        #
        # Previous position was near the end and WMP now changed
        # to stopped/ready.
        #
        # This specifically handles:
        #
        #   previous:
        #       state=3
        #       position=239.0
        #
        #   current:
        #       state=1
        #       position=0
        # --------------------------------------------------------

        if (
            previous_seen_playing
            and previous_duration > 1.0
            and previous_position
                >= previous_duration - 2.0
            and state in (1, 9, 10)
        ):

            logger.info(
                "WMP EOF detected by position/state transition: "
                "previous_position=%.2f / %.2f, "
                "current_position=%.2f, state=%s",
                previous_position,
                previous_duration,
                current_position,
                state,
            )

            self._media_end_handled = True
            self._seen_playing = False

            self._handle_media_ended()

            return

        # --------------------------------------------------------
        # Method 3:
        #
        # Use accumulated actual playing time.
        #
        # This is the important fallback for WMP codecs that reset
        # currentPosition to 0 immediately when playback ends.
        # --------------------------------------------------------

        if (
            previous_seen_playing
            and previous_duration > 1.0
            and self._playing_time
                >= previous_duration - 1.5
            and state in (1, 9, 10)
        ):

            logger.info(
                "WMP EOF detected by elapsed playback time: "
                "played_time=%.2f / %.2f, "
                "position=%.2f, state=%s",
                self._playing_time,
                previous_duration,
                current_position,
                state,
            )

            self._media_end_handled = True
            self._seen_playing = False

            self._handle_media_ended()

            return

    # ============================================================
    # Handle Media End
    # ============================================================

    def _handle_media_ended(self):
        """
        Handle the end of the current song.
        """

        global current_index
        global playlist_auto_advance
        global random_play_enabled
        global random_play_running
        global random_play_played

        # --------------------------------------------------------
        # Random playback
        # --------------------------------------------------------

        if random_play_enabled and random_play_running:

            random_play_played += 1

            logger.info(
                "Random playback progress: %d / %d",
                random_play_played,
                random_play_count,
            )

            # More random songs remain.
            if current_index + 1 < len(music_playlist):

                current_index += 1

                next_song = Path(
                    music_playlist[current_index]
                )

                logger.info(
                    "Random playback next song: %s",
                    next_song,
                )

                self._play_path(next_song)

            else:

                logger.info(
                    "Random playback finished: %d songs.",
                    random_play_played,
                )

                random_play_enabled = False
                random_play_running = False
                playlist_auto_advance = False

                current_index = -1

            return

        # --------------------------------------------------------
        # Normal playlist automatic playback
        # --------------------------------------------------------

        if playlist_auto_advance:

            if current_index + 1 < len(music_playlist):

                current_index += 1

                next_song = Path(
                    music_playlist[current_index]
                )

                logger.info(
                    "Playlist next song: %s",
                    next_song,
                )

                self._play_path(next_song)

            else:

                logger.info(
                    "Playlist finished."
                )

                playlist_auto_advance = False
                current_index = -1

    # ============================================================
    # Play
    # ============================================================

    def _play_path(self, path: Path):
        """
        Play a local file through Windows Media Player.
        """

        if self.wmp is None:

            raise RuntimeError(
                "Windows Media Player is not initialized."
            )

        path = path.resolve()

        if not path.exists():

            raise FileNotFoundError(
                f"Music file does not exist: {path}"
            )

        if random_play_enabled and random_play_running:
            logger.info(
                "Random playback: %d / %d - Playing: %s",
                current_index + 1,
                random_play_count,
                path,
            )
        else:
            logger.info(
                "Playing: %s",
                path,
            )

        # --------------------------------------------------------
        # Reset playback tracking for the new song
        # --------------------------------------------------------

        self._media_end_handled = False
        self._seen_playing = False

        self._last_position = 0.0
        self._last_duration = 0.0

        self._last_logged_state = None

        self._playing_time = 0.0
        self._last_check_time = time.monotonic()
        self._last_state = None

        # --------------------------------------------------------
        # Create media only ONCE
        # --------------------------------------------------------

        logger.info(
            "WMP: creating media object..."
        )

        media = self.wmp.newMedia(
            str(path)
        )

        logger.info(
            "WMP: media object created."
        )

        # Get duration from this same media object.
        try:

            media_duration = float(
                media.duration
            )

            if media_duration > 0:

                self._last_duration = (
                    media_duration
                )

        except Exception:

            self._last_duration = 0.0

        logger.info(
            "WMP: media duration = %.2f seconds",
            self._last_duration,
        )

        # --------------------------------------------------------
        # Assign media
        # --------------------------------------------------------

        self.wmp.currentMedia = media

        logger.info(
            "WMP: currentMedia assigned."
        )

        # --------------------------------------------------------
        # Start playback
        # --------------------------------------------------------

        self.wmp.controls.play()

        logger.info(
            "WMP: controls.play() returned."
        )

    # ============================================================
    # Public Play
    # ============================================================

    def play(self, path: Path):

        self.command_queue.put(
            (
                self._play_path,
                (path,),
            )
        )

    # ============================================================
    # Pause
    # ============================================================

    def pause(self):

        def _pause():

            if self.wmp is not None:

                self.wmp.controls.pause()

        self.command_queue.put(
            (
                _pause,
                (),
            )
        )

    # ============================================================
    # Resume
    # ============================================================

    def resume(self):

        def _resume():

            if self.wmp is not None:

                self.wmp.controls.play()

        self.command_queue.put(
            (
                _resume,
                (),
            )
        )

    # ============================================================
    # Stop
    # ============================================================

    def stop(self):

        def _stop():

            if self.wmp is not None:

                self.wmp.controls.stop()

        self.command_queue.put(
            (
                _stop,
                (),
            )
        )

    # ============================================================
    # Close
    # ============================================================

    def close(self):

        self.running = False


# ============================================================
# Global WMP Controller
# ============================================================

wmp_controller = WMPController()


# ============================================================
# MCP Tools
# ============================================================

@mcp.tool()
def play_song(song_path: str) -> dict:
    """
    Play a local music file on the LOCAL WINDOWS COMPUTER.

    The path can be either:
    - an absolute Windows path
    - a path relative to the configured music directory

    Example:
        周杰伦\\晴天.mp3
    """

    global music_playlist
    global current_index
    global playlist_auto_advance
    global random_play_enabled
    global random_play_running

    try:
        path = _resolve_music_path(song_path)

        if not path.exists():
            return {
                "success": False,
                "message": f"Music file not found: {path}",
            }

        if not path.is_file():
            return {
                "success": False,
                "message": f"Not a music file: {path}",
            }

        if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            return {
                "success": False,
                "message": f"Unsupported music format: {path.suffix}",
            }

        # Manual play cancels random playback mode.
        random_play_enabled = False
        random_play_running = False
        playlist_auto_advance = False

        music_playlist = [str(path)]
        current_index = 0

        wmp_controller.play(path)

        return {
            "success": True,
            "message": f"Playing: {path.name}",
            "song": _song_info(path),
        }

    except Exception as e:
        logger.exception("play_song failed")

        return {
            "success": False,
            "message": str(e),
        }


@mcp.tool()
def play_folder(
    directory: str = DEFAULT_MUSIC_DIRECTORY,
) -> dict:
    """
    Scan a local music folder recursively and play all songs sequentially.

    Songs inside subdirectories are included.
    """

    global music_playlist
    global current_index
    global playlist_directory
    global playlist_auto_advance
    global random_play_enabled
    global random_play_running

    try:

        songs = scan_music_directory(directory)

        if not songs:
            return {
                "success": False,
                "message": "No supported music files found.",
            }

        music_playlist = [
            str(path.resolve())
            for path in songs
        ]

        current_index = 0
        playlist_directory = directory

        playlist_auto_advance = True

        random_play_enabled = False
        random_play_running = False

        first_song = songs[0]

        wmp_controller.play(first_song)

        return {
            "success": True,
            "message": (
                f"Found {len(songs)} songs. "
                f"Started playlist."
            ),
            "total_songs": len(songs),
            "current_song": _song_info(first_song),
        }

    except Exception as e:
        logger.exception("play_folder failed")

        return {
            "success": False,
            "message": str(e),
        }


@mcp.tool()
def random_play(
    count: int = 10,
    directory: str = DEFAULT_MUSIC_DIRECTORY,
) -> dict:
    """
    Randomly select and play a specified number of unique songs.

    Default:
        10 songs

    Rules:
        - Songs are selected randomly.
        - No duplicate songs within the same random playback.
        - Songs are played sequentially.
        - Playback automatically advances when each song ends.
        - Playback stops automatically after the requested number
          of songs have finished.
        - If count is larger than the available number of songs,
          all available songs are played.
    """

    global music_playlist
    global current_index
    global playlist_directory
    global playlist_auto_advance
    global random_play_enabled
    global random_play_running
    global random_play_count
    global random_play_played

    try:

        # Validate count.
        if count < 1:
            return {
                "success": False,
                "message": "count must be at least 1.",
            }

        songs = scan_music_directory(directory)

        if not songs:
            return {
                "success": False,
                "message": "No supported music files found.",
            }

        # If requested count exceeds total songs,
        # simply play every available song once.
        actual_count = min(count, len(songs))

        # random.sample() guarantees no duplicates.
        selected_songs = random.sample(
            songs,
            actual_count,
        )

        music_playlist = [
            str(path.resolve())
            for path in selected_songs
        ]

        current_index = 0

        playlist_directory = directory

        # Enable random playback mode.
        random_play_enabled = True
        random_play_running = True

        random_play_count = actual_count
        random_play_played = 0

        # Disable normal playlist mode.
        playlist_auto_advance = False

        first_song = selected_songs[0]

        logger.info(
            "Starting random playback: %d songs.",
            actual_count,
        )

        wmp_controller.play(first_song)

        return {
            "success": True,
            "message": (
                f"Randomly selected {actual_count} songs. "
                f"Playback will stop after all selected songs finish."
            ),
            "requested_count": count,
            "actual_count": actual_count,
            "current_song": _song_info(first_song),
            "songs": [
                _song_info(path)
                for path in selected_songs
            ],
        }

    except Exception as e:
        logger.exception("random_play failed")

        return {
            "success": False,
            "message": str(e),
        }


@mcp.tool()
def stop_random_play() -> dict:
    """
    Stop automatic random-play progression.

    This disables random playback mode.
    The currently playing song is not forcibly stopped.
    Use stop_music() if the music itself should also stop.
    """

    global random_play_enabled
    global random_play_running
    global playlist_auto_advance

    random_play_enabled = False
    random_play_running = False
    playlist_auto_advance = False

    logger.info("Random playback mode disabled.")

    return {
        "success": True,
        "message": "Random playback mode stopped.",
    }


@mcp.tool()
def play_random_song(
    directory: str = DEFAULT_MUSIC_DIRECTORY,
) -> dict:
    """
    Randomly select one local music file and play it.
    """

    global music_playlist
    global current_index
    global playlist_auto_advance
    global random_play_enabled
    global random_play_running

    try:

        songs = scan_music_directory(directory)

        if not songs:
            return {
                "success": False,
                "message": "No supported music files found.",
            }

        song = random.choice(songs)

        music_playlist = [str(song)]
        current_index = 0

        playlist_auto_advance = False

        random_play_enabled = False
        random_play_running = False

        wmp_controller.play(song)

        return {
            "success": True,
            "message": f"Randomly selected: {song.name}",
            "song": _song_info(song),
        }

    except Exception as e:
        logger.exception("play_random_song failed")

        return {
            "success": False,
            "message": str(e),
        }


@mcp.tool()
def pause_music() -> dict:
    """
    Pause the currently playing music on the LOCAL WINDOWS COMPUTER.
    """

    try:

        wmp_controller.pause()

        return {
            "success": True,
            "message": "Music paused.",
        }

    except Exception as e:
        logger.exception("pause_music failed")

        return {
            "success": False,
            "message": str(e),
        }


@mcp.tool()
def resume_music() -> dict:
    """
    Resume paused music on the LOCAL WINDOWS COMPUTER.
    """

    try:

        wmp_controller.resume()

        return {
            "success": True,
            "message": "Music resumed.",
        }

    except Exception as e:
        logger.exception("resume_music failed")

        return {
            "success": False,
            "message": str(e),
        }


@mcp.tool()
def stop_music() -> dict:
    """
    Stop music playback on the LOCAL WINDOWS COMPUTER.
    """

    global playlist_auto_advance
    global random_play_enabled
    global random_play_running

    try:

        playlist_auto_advance = False

        random_play_enabled = False
        random_play_running = False

        wmp_controller.stop()

        return {
            "success": True,
            "message": "Music stopped.",
        }

    except Exception as e:
        logger.exception("stop_music failed")

        return {
            "success": False,
            "message": str(e),
        }

@mcp.tool()
def play_next_song() -> dict:
    """
    Play the next song in the current playlist.

    Behavior:
    - During random playback:
        Move to the next randomly selected song while keeping
        random playback mode active.
        The current random playback position is updated.
        random_play_played is not changed because manually skipping
        a song does not mean that the song has finished playing.

    - During normal playback:
        Keep the existing playlist behavior.
    """

    global current_index
    global playlist_auto_advance
    global random_play_enabled
    global random_play_running

    try:

        if not music_playlist:
            return {
                "success": False,
                "message": "No playlist is currently loaded.",
            }

        # ========================================================
        # Random playback mode
        # ========================================================

        if random_play_enabled and random_play_running:

            # Check whether the current song is already the last
            # song in the random playback sequence.
            if current_index + 1 >= len(music_playlist):
                return {
                    "success": False,
                    "message": "Already at the last random song.",
                    "index": current_index + 1,
                    "total": random_play_count,
                }

            # Move to the next random song.
            current_index += 1

            path = Path(
                music_playlist[current_index]
            ).resolve()

            # Keep random playback enabled.
            random_play_enabled = True
            random_play_running = True

            # Normal automatic playlist mode remains disabled.
            playlist_auto_advance = False

            logger.info(
                "Manual random next: %d / %d - Playing: %s",
                current_index + 1,
                random_play_count,
                path,
            )

            wmp_controller.play(path)

            return {
                "success": True,
                "message": (
                    f"Playing next random song: "
                    f"{path.name} "
                    f"({current_index + 1} / {random_play_count})"
                ),
                "song": _song_info(path),

                # Current position in the random sequence
                "index": current_index + 1,

                # Total random songs
                "total": random_play_count,

                # Random playback state
                "random_play": True,
                "random_play_count": random_play_count,

                # Number of songs that actually finished
                # playing is NOT changed by manual skipping.
                "random_play_played": random_play_played,
            }

        # ========================================================
        # Normal playlist mode
        # ========================================================

        # Disable random playback if it is not currently active.
        random_play_enabled = False
        random_play_running = False

        playlist_auto_advance = True

        if current_index + 1 >= len(music_playlist):
            return {
                "success": False,
                "message": "Already at the last song.",
            }

        current_index += 1

        path = Path(
            music_playlist[current_index]
        ).resolve()

        logger.info(
            "Manual next: %d / %d - Playing: %s",
            current_index + 1,
            len(music_playlist),
            path,
        )

        wmp_controller.play(path)

        return {
            "success": True,
            "message": f"Playing next: {path.name}",
            "song": _song_info(path),
            "index": current_index + 1,
            "total": len(music_playlist),
            "random_play": False,
        }

    except Exception as e:

        logger.exception(
            "play_next_song failed"
        )

        return {
            "success": False,
            "message": str(e),
        }


@mcp.tool()
def play_previous_song() -> dict:
    """
    Play the previous song in the current playlist.

    Behavior:
    - During random playback:
        Move to the previous randomly selected song while keeping
        random playback mode active.
        The current random playback position is updated.
        random_play_played is not changed.

    - During normal playback:
        Keep the existing playlist behavior.
    """

    global current_index
    global playlist_auto_advance
    global random_play_enabled
    global random_play_running

    try:

        if not music_playlist:
            return {
                "success": False,
                "message": "No playlist is currently loaded.",
            }

        # ========================================================
        # Random playback mode
        # ========================================================

        if random_play_enabled and random_play_running:

            # Check whether the current song is already the first
            # song in the random playback sequence.
            if current_index <= 0:
                return {
                    "success": False,
                    "message": "Already at the first random song.",
                    "index": current_index + 1,
                    "total": random_play_count,
                }

            # Move to the previous random song.
            current_index -= 1

            path = Path(
                music_playlist[current_index]
            ).resolve()

            # Keep random playback enabled.
            random_play_enabled = True
            random_play_running = True

            # Normal automatic playlist mode remains disabled.
            playlist_auto_advance = False

            logger.info(
                "Manual random previous: %d / %d - Playing: %s",
                current_index + 1,
                random_play_count,
                path,
            )

            wmp_controller.play(path)

            return {
                "success": True,
                "message": (
                    f"Playing previous random song: "
                    f"{path.name} "
                    f"({current_index + 1} / {random_play_count})"
                ),
                "song": _song_info(path),

                # Current position in the random sequence
                "index": current_index + 1,

                # Total random songs
                "total": random_play_count,

                # Random playback state
                "random_play": True,
                "random_play_count": random_play_count,

                # Number of songs that actually finished
                # playing is not changed by manual navigation.
                "random_play_played": random_play_played,
            }

        # ========================================================
        # Normal playlist mode
        # ========================================================

        # Disable random playback if it is not currently active.
        random_play_enabled = False
        random_play_running = False

        playlist_auto_advance = True

        if current_index <= 0:
            return {
                "success": False,
                "message": "Already at the first song.",
            }

        current_index -= 1

        path = Path(
            music_playlist[current_index]
        ).resolve()

        logger.info(
            "Manual previous: %d / %d - Playing: %s",
            current_index + 1,
            len(music_playlist),
            path,
        )

        wmp_controller.play(path)

        return {
            "success": True,
            "message": f"Playing previous: {path.name}",
            "song": _song_info(path),
            "index": current_index + 1,
            "total": len(music_playlist),
            "random_play": False,
        }
    except Exception as e:

        logger.exception(
            "play_previous_song failed"
        )

        return {
            "success": False,
            "message": str(e),
        }

@mcp.tool()
def get_current_music_info() -> dict:
    """
    Get information about the currently playing music.

    Returns:
    - index: current song number, starting from 1
    - total: total number of songs
    - random_play: whether random playback is active
    - random_play_count: total number of songs planned for this random playback
    - random_play_played: number of songs that have already finished playing

    Note:
    random_play_played represents the number of completed songs,
    not the current song number.

    The current song number is determined by current_index + 1.
    """

    try:
        # ========================================================
        # No music playlist
        # ========================================================

        if not music_playlist:
            return {
                "success": True,
                "playing": False,
                "message": "No music is currently selected.",
            }

        # ========================================================
        # Invalid current index
        # ========================================================

        if current_index < 0 or current_index >= len(music_playlist):
            return {
                "success": True,
                "playing": False,
                "message": "No active playlist item.",
            }

        # ========================================================
        # Get the current song
        # ========================================================

        path = Path(
            music_playlist[current_index]
        ).resolve()

        # ========================================================
        # Current song number
        #
        # current_index:
        #   0 -> Song 1
        #   1 -> Song 2
        #   2 -> Song 3
        # ========================================================

        current_number = current_index + 1

        # ========================================================
        # Determine the total number of songs
        #
        # Random playback:
        #   Use the number of songs planned for this random playback.
        #
        # Normal playback:
        #   Use the number of songs in the current playlist.
        # ========================================================

        if random_play_enabled and random_play_running:
            total_number = random_play_count
        else:
            total_number = len(music_playlist)

        # ========================================================
        # Return current music information
        # ========================================================

        return {
            "success": True,

            # A song is currently selected and playing
            "playing": True,

            # Detailed information about the current song
            "song": _song_info(path),

            # Current song number, starting from 1
            "index": current_number,

            # Total number of songs
            "total": total_number,

            # ====================================================
            # Random playback information
            # ====================================================

            # Whether random playback is currently active
            "random_play": random_play_enabled,

            # Total number of songs planned for this random playback
            "random_play_count": random_play_count,

            # Number of songs that have already finished playing
            #
            # Example:
            # First song is playing -> 0
            # First song finished, second song is playing -> 1
            # Second song finished -> 2
            "random_play_played": random_play_played,

            # ====================================================
            # Explicitly distinguish the current song number
            # from the number of completed songs
            # ====================================================

            # Current song number in the random playback sequence
            "current_random_index": (
                current_number
                if random_play_enabled and random_play_running
                else None
            ),

            # Total number of songs in the random playback sequence
            "current_random_total": (
                random_play_count
                if random_play_enabled and random_play_running
                else None
            ),
        }

    except Exception as e:

        logger.exception(
            "get_current_music_info failed"
        )

        return {
            "success": False,
            "message": str(e),
        }



@mcp.tool()
def list_local_songs(
    directory: str = DEFAULT_MUSIC_DIRECTORY,
) -> dict:
    """
    List all supported local music files recursively.

    Relative paths are returned so songs inside subdirectories
    can be uniquely identified.

    Example:
        周杰伦\\晴天.mp3
        林俊杰\\江南.mp3
    """

    try:

        songs = scan_music_directory(directory)

        return {
            "success": True,
            "directory": str(
                Path(directory).resolve()
            ),
            "total": len(songs),
            "songs": [
                _get_relative_music_path(path)
                for path in songs
            ],
        }

    except Exception as e:
        logger.exception("list_local_songs failed")

        return {
            "success": False,
            "message": str(e),
        }


@mcp.tool()
def scan_local_music_folder(
    directory: str = DEFAULT_MUSIC_DIRECTORY,
) -> dict:
    """
    Scan a local music folder recursively and return detailed
    information about all supported music files.
    """

    try:

        songs = scan_music_directory(directory)

        return {
            "success": True,
            "directory": str(
                Path(directory).resolve()
            ),
            "total": len(songs),
            "songs": [
                _song_info(path)
                for path in songs
            ],
        }

    except Exception as e:
        logger.exception("scan_local_music_folder failed")

        return {
            "success": False,
            "message": str(e),
        }


# ============================================================
# Windows Master Volume
# ============================================================

@mcp.tool()
def set_windows_system_volume(
    volume_percentage: int,
) -> dict:
    """
    Set the master volume of the LOCAL WINDOWS COMPUTER.

    volume_percentage:
        0  = mute
        100 = maximum volume
    """

    if volume_percentage < 0 or volume_percentage > 100:
        return {
            "success": False,
            "message": "Volume must be between 0 and 100.",
        }

    result = {
        "success": False,
        "error": None,
    }

    def change_volume():
        try:

            import comtypes
            from pycaw.pycaw import AudioUtilities

            # COM must be initialized on the thread that uses pycaw.
            comtypes.CoInitialize()

            try:

                devices = AudioUtilities.GetSpeakers()

                volume_interface = devices.EndpointVolume

                volume_interface.SetMasterVolumeLevelScalar(
                    volume_percentage / 100.0,
                    None,
                )

                result["success"] = True

            finally:

                comtypes.CoUninitialize()

        except Exception as e:
            result["error"] = e

    thread = threading.Thread(
        target=change_volume,
        daemon=True,
        name="WindowsVolumeController",
    )

    thread.start()
    thread.join()

    if result["success"]:

        return {
            "success": True,
            "message": (
                f"Windows master volume set to "
                f"{volume_percentage}%."
            ),
            "volume": volume_percentage,
        }

    return {
        "success": False,
        "message": str(result["error"]),
    }


# ============================================================
# Server Entry Point
# ============================================================

if __name__ == "__main__":

    logger.info(
        "Starting Local Windows Music Player MCP Server..."
    )

    logger.info(
        "Default music directory: %s",
        DEFAULT_MUSIC_DIRECTORY,
    )

    # MCP must use stdout for its stdio protocol.
    # Do not print anything manually to stdout.
    mcp.run(transport="stdio")
