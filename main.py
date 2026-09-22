import os
import sys
import json
import time
import queue
import re
from difflib import SequenceMatcher

import sounddevice as sd
import vosk
import win32com.client


# ============================================================
# HERITAGE AI
# Voice Presentation Assistant
# ============================================================

APP_NAME = "HERITAGE AI"


# ============================================================
# PATHS
# ============================================================

def get_base_path():

    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)

    return os.path.dirname(
        os.path.abspath(__file__)
    )


BASE_PATH = get_base_path()

PRESENTATION_FILE = os.path.join(
    BASE_PATH,
    "HeritageAI.pptx"
)

MODEL_FOLDER = "vosk-model-small-en-us-0.15"

# Possible model locations.
#
# Development:
#   HeritageAI\_internal\vosk-model-small-en-us-0.15
#   HeritageAI\vosk-model-small-en-us-0.15
#
# One-file EXE:
#   temporary_extract\vosk-model-small-en-us-0.15
#   temporary_extract\_internal\vosk-model-small-en-us-0.15
#
# We check all of them.

if getattr(sys, "frozen", False):

    FROZEN_BASE = getattr(
        sys,
        "_MEIPASS",
        os.path.dirname(sys.executable)
    )

    MODEL_PATH_INTERNAL = os.path.join(
        FROZEN_BASE,
        "_internal",
        MODEL_FOLDER
    )

    MODEL_PATH_PACKAGED = os.path.join(
        FROZEN_BASE,
        MODEL_FOLDER
    )

else:

    MODEL_PATH_INTERNAL = os.path.join(
        BASE_PATH,
        "_internal",
        MODEL_FOLDER
    )

    MODEL_PATH_PACKAGED = os.path.join(
        BASE_PATH,
        MODEL_FOLDER
    )


# ============================================================
# GLOBAL VARIABLES
# ============================================================

model = None
recognizer = None

powerpoint = None
presentation = None

audio_queue = queue.Queue()

sample_rate = 44100
microphone_index = None

# Cached presentation intelligence index.
slide_index = []


# ============================================================
# NUMBER WORDS
# ============================================================

NUMBER_WORDS = {

    "zero": 0,

    "one": 1,
    "won": 1,
    "wun": 1,
    "wan": 1,

    "two": 2,
    "too": 2,
    "to": 2,
    "twu": 2,

    "three": 3,
    "tree": 3,
    "thre": 3,
    "free": 3,
    "tri": 3,
    "tiri": 3,

    "four": 4,
    "for": 4,
    "fore": 4,
    "fo": 4,
    "faw": 4,

    "five": 5,
    "fiv": 5,
    "faiv": 5,

    "six": 6,
    "sicks": 6,
    "sik": 6,

    "seven": 7,
    "sevin": 7,
    "sevan": 7,

    "eight": 8,
    "ate": 8,
    "ait": 8,
    "eit": 8,

    "nine": 9,
    "nain": 9,
    "neen": 9,

    "ten": 10,
    "tin": 10,
    "then": 10,

    "eleven": 11,
    "leven": 11,

    "twelve": 12,
    "twelv": 12,

    "thirteen": 13,
    "thirteenth": 13,

    "fourteen": 14,
    "forteen": 14,

    "fifteen": 15,
    "fiveteen": 15,

    "sixteen": 16,
    "sixten": 16,

    "seventeen": 17,
    "seventen": 17,
}


# ============================================================
# COMMON SPEECH RECOGNITION CORRECTIONS
# ============================================================

SPEECH_CORRECTIONS = {

    # Heritage
    "heritagee": "heritage",
    "heritagey": "heritage",
    "heritages": "heritage",
    "heritageai": "heritage",
    "heritage a i": "heritage",
    "heritage": "heritage",
    "her itage": "heritage",
    "heratage": "heritage",
    "heretage": "heritage",
    "heritag": "heritage",
    "heritash": "heritage",
    "heritaj": "heritage",
    "heritige": "heritage",
    "heritigee": "heritage",
    "herigate": "heritage",
    "heritace": "heritage",
    "heritich": "heritage",

    # Numbers
    "won": "one",
    "wun": "one",
    "wan": "one",

    "too": "two",
    "to": "two",
    "twu": "two",

    "tree": "three",
    "thre": "three",
    "free": "three",
    "tri": "three",
    "tiri": "three",

    "for": "four",
    "fore": "four",
    "fo": "four",
    "faw": "four",

    "fiv": "five",
    "faiv": "five",

    "sicks": "six",
    "sik": "six",

    "sevin": "seven",
    "sevan": "seven",

    "ate": "eight",
    "ait": "eight",
    "eit": "eight",

    "nain": "nine",
    "neen": "nine",

    "tin": "ten",
    "then": "ten",

    "leven": "eleven",

    "twelv": "twelve",

    "thirteenth": "thirteen",
    "forteen": "fourteen",
    "fiveteen": "fifteen",
    "sixten": "sixteen",
    "seventen": "seventeen",
}


# ============================================================
# COMMAND PHRASES
# ============================================================

NEXT_PHRASES = {
    "next",
    "go next",
    "go forward",
    "move forward",
    "move to next",
    "move to the next",
    "next one",
    "forward",
    "forward one",
    "advance",
    "go ahead",
    "continue",
    "next slide",
    "go to next",
    "go to the next",
}

PREVIOUS_PHRASES = {
    "previous",
    "go previous",
    "go back",
    "move back",
    "move backward",
    "back",
    "back one",
    "previous one",
    "go to previous",
    "go to the previous",
    "previous slide",
}

START_PHRASES = {
    "start",
    "start presentation",
    "begin",
    "begin presentation",
    "open presentation",
    "play presentation",
    "start slideshow",
    "begin slideshow",
    "show presentation",
    "launch presentation",
}

STOP_PHRASES = {
    "stop",
    "stop presentation",
    "end presentation",
    "end slideshow",
    "exit presentation",
    "close presentation",
    "stop slideshow",
}


# ============================================================
# NORMALIZE TEXT
# ============================================================

def normalize_text(text):

    if not text:
        return ""

    text = text.lower().strip()

    text = re.sub(
        r"[^a-z0-9\s]",
        " ",
        text
    )

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


# ============================================================
# SIMILARITY
# ============================================================

def similarity(a, b):

    return SequenceMatcher(
        None,
        a,
        b
    ).ratio()


# ============================================================
# CORRECT SPEECH
# ============================================================

def correct_speech_words(text):

    text = normalize_text(text)

    if not text:
        return ""

    # Handle multi-word Heritage variations first.

    multi_word_corrections = {
        "heritage ai": "heritage",
        "heritage a i": "heritage",
        "her itage": "heritage",
    }

    for wrong, correct in multi_word_corrections.items():

        text = text.replace(
            wrong,
            correct
        )

    words = text.split()

    corrected = []

    for word in words:

        if word in SPEECH_CORRECTIONS:

            corrected.append(
                SPEECH_CORRECTIONS[word]
            )

        else:

            corrected.append(word)

    return " ".join(corrected)


# ============================================================
# FIND MICROPHONE
# ============================================================

def find_microphone():

    global microphone_index
    global sample_rate

    try:

        devices = sd.query_devices()

        preferred_index = 1

        if preferred_index < len(devices):

            device = devices[
                preferred_index
            ]

            if device[
                "max_input_channels"
            ] > 0:

                microphone_index = (
                    preferred_index
                )

        if microphone_index is None:

            for i, device in enumerate(
                devices
            ):

                if device[
                    "max_input_channels"
                ] > 0:

                    microphone_index = i

                    break

        if microphone_index is None:

            print()
            print(
                "HeritageAI: "
                "No microphone was found."
            )

            return False

        print()
        print(
            "HeritageAI: "
            "Using microphone:",
            microphone_index
        )

        sample_rate = 44100

        print(
            "HeritageAI: "
            "Using sample rate:",
            sample_rate
        )

        return True

    except Exception as error:

        print()
        print(
            "HeritageAI: "
            "Could not access microphone information."
        )

        print(
            "Error:",
            error
        )

        return False


# ============================================================
# FIND POWERPOINT APPLICATION
# ============================================================

def find_powerpoint_window():

    global powerpoint

    try:

        if powerpoint is None:

            powerpoint = (
                win32com.client.GetActiveObject(
                    "PowerPoint.Application"
                )
            )

        return powerpoint

    except Exception:

        return None


# ============================================================
# FOCUS POWERPOINT
# ============================================================

def focus_powerpoint():

    try:

        import win32gui
        import win32con

        windows = []

        def enum_windows(
            hwnd,
            extra
        ):

            try:

                title = (
                    win32gui.GetWindowText(
                        hwnd
                    )
                )

                if not title:
                    return

                title_lower = (
                    title.lower()
                )

                if (
                    "powerpoint slide show"
                    in title_lower
                    or
                    "powerpoint"
                    in title_lower
                ):

                    if win32gui.IsWindowVisible(
                        hwnd
                    ):

                        extra.append(hwnd)

            except Exception:

                pass

        win32gui.EnumWindows(
            enum_windows,
            windows
        )

        # Prefer slideshow window.

        for hwnd in windows:

            try:

                title = (
                    win32gui.GetWindowText(
                        hwnd
                    )
                )

                if (
                    "powerpoint slide show"
                    in title.lower()
                ):

                    win32gui.ShowWindow(
                        hwnd,
                        win32con.SW_RESTORE
                    )

                    win32gui.SetForegroundWindow(
                        hwnd
                    )

                    return

            except Exception:

                pass

        # Otherwise focus normal PowerPoint.

        for hwnd in windows:

            try:

                win32gui.ShowWindow(
                    hwnd,
                    win32con.SW_RESTORE
                )

                win32gui.SetForegroundWindow(
                    hwnd
                )

                return

            except Exception:

                pass

    except Exception:

        pass


# ============================================================
# GET ACTIVE PRESENTATION
# ============================================================

def get_active_presentation():

    global powerpoint

    try:

        if powerpoint is None:

            powerpoint = (
                win32com.client.GetActiveObject(
                    "PowerPoint.Application"
                )
            )

        # Currently active presentation.

        try:

            active = (
                powerpoint.ActivePresentation
            )

            if active is not None:

                return active

        except Exception:

            pass

        # If only one presentation exists.

        try:

            if (
                powerpoint.Presentations.Count
                == 1
            ):

                return (
                    powerpoint.Presentations.Item(
                        1
                    )
                )

        except Exception:

            pass

        # Final fallback.

        try:

            if (
                powerpoint.Presentations.Count
                > 0
            ):

                return (
                    powerpoint.Presentations.Item(
                        1
                    )
                )

        except Exception:

            pass

    except Exception:

        pass

    return None


# ============================================================
# CONNECT TO POWERPOINT
# ============================================================

def connect_powerpoint():

    global powerpoint
    global presentation

    print()
    print(
        "HeritageAI: "
        "Connecting to PowerPoint..."
    )

    try:

        # ----------------------------------------------------
        # Connect to running PowerPoint.
        # ----------------------------------------------------

        try:

            powerpoint = (
                win32com.client.GetActiveObject(
                    "PowerPoint.Application"
                )
            )

            print(
                "HeritageAI: "
                "PowerPoint application found."
            )

        except Exception:

            powerpoint = (
                win32com.client.Dispatch(
                    "PowerPoint.Application"
                )
            )

            print(
                "HeritageAI: "
                "PowerPoint application started."
            )

        # ----------------------------------------------------
        # Find presentation.
        # ----------------------------------------------------

        presentation = (
            get_active_presentation()
        )

        # ----------------------------------------------------
        # Existing presentation.
        # ----------------------------------------------------

        if presentation is not None:

            try:

                presentation_name = (
                    presentation.Name
                )

            except Exception:

                presentation_name = (
                    "Unknown presentation"
                )

            print()
            print(
                "HeritageAI: "
                "Active presentation found:"
            )

            print(
                "         ",
                presentation_name
            )

        # ----------------------------------------------------
        # Fallback presentation.
        # ----------------------------------------------------

        else:

            if not os.path.exists(
                PRESENTATION_FILE
            ):

                print()
                print(
                    "HeritageAI: "
                    "No PowerPoint presentation is open."
                )

                print()
                print(
                    "HeritageAI also could not find "
                    "the fallback presentation:"
                )

                print(
                    PRESENTATION_FILE
                )

                return False

            print()
            print(
                "HeritageAI: "
                "No presentation was open."
            )

            print(
                "HeritageAI: "
                "Opening fallback presentation..."
            )

            presentation = (
                powerpoint.Presentations.Open(
                    PRESENTATION_FILE,
                    False,
                    False,
                    True
                )
            )

            time.sleep(1)

        if presentation is None:

            print(
                "HeritageAI: "
                "Presentation unavailable."
            )

            return False

        # ----------------------------------------------------
        # Presentation information.
        # ----------------------------------------------------

        try:

            name = (
                presentation.Name
            )

        except Exception:

            name = "Unknown"

        try:

            slide_count = (
                presentation.Slides.Count
            )

        except Exception:

            slide_count = 0

        print()
        print(
            "HeritageAI: "
            "Using presentation:"
        )

        print(
            "         ",
            name
        )

        print(
            "HeritageAI: "
            "Number of slides:",
            slide_count
        )

        print()
        print(
            "HeritageAI: "
            "Connected to PowerPoint."
        )

        build_slide_index()

        focus_powerpoint()

        return True

    except Exception as error:

        print()
        print(
            "HeritageAI: "
            "Could not connect "
            "to PowerPoint."
        )

        print(
            "Error:",
            error
        )

        return False


# ============================================================
# GET SLIDESHOW VIEW
# ============================================================

def get_slideshow_view():

    global presentation

    if presentation is None:

        return None

    try:

        window = (
            presentation.SlideShowWindow
        )

        if window is None:

            return None

        view = window.View

        if view is None:

            return None

        return view

    except Exception:

        return None


# ============================================================
# CHECK SLIDESHOW
# ============================================================

def slideshow_is_running():

    try:

        return (
            get_slideshow_view()
            is not None
        )

    except Exception:

        return False


# ============================================================
# START PRESENTATION
# ============================================================

def start_presentation():

    global presentation
    global powerpoint

    print()
    print(
        "HeritageAI: "
        "Starting presentation..."
    )

    try:

        if presentation is None:

            if not connect_powerpoint():

                return False

        if slideshow_is_running():

            print(
                "HeritageAI: "
                "Slideshow is already running."
            )

            focus_powerpoint()

            return True

        settings = (
            presentation.SlideShowSettings
        )

        if settings is None:

            print(
                "HeritageAI: "
                "SlideShowSettings unavailable."
            )

            return False

        settings.StartingSlide = 1

        settings.EndingSlide = (
            presentation.Slides.Count
        )

        settings.Run()

        time.sleep(1)

        view = (
            get_slideshow_view()
        )

        if view is None:

            print()
            print(
                "HeritageAI: "
                "Presentation did not start."
            )

            return False

        print(
            "HeritageAI: "
            "Presentation started."
        )

        focus_powerpoint()

        return True

    except Exception as error:

        print()
        print(
            "HeritageAI: "
            "Could not start presentation."
        )

        print(
            "Error:",
            error
        )

        return False


# ============================================================
# NEXT SLIDE
# ============================================================

def next_slide():

    if not slideshow_is_running():

        print(
            "HeritageAI: "
            "Slideshow is not running."
        )

        print(
            "HeritageAI: "
            "Starting presentation first..."
        )

        if not start_presentation():

            return

    try:

        view = (
            get_slideshow_view()
        )

        if view is None:

            return

        view.Next()

        print(
            "HeritageAI: "
            "Next slide."
        )

        focus_powerpoint()

    except Exception as error:

        print(
            "HeritageAI: "
            "Could not go to next slide."
        )

        print(
            "Error:",
            error
        )


# ============================================================
# PREVIOUS SLIDE
# ============================================================

def previous_slide():

    if not slideshow_is_running():

        print(
            "HeritageAI: "
            "Slideshow is not running."
        )

        print(
            "HeritageAI: "
            "Starting presentation first..."
        )

        if not start_presentation():

            return

    try:

        view = (
            get_slideshow_view()
        )

        if view is None:

            return

        view.Previous()

        print(
            "HeritageAI: "
            "Previous slide."
        )

        focus_powerpoint()

    except Exception as error:

        print(
            "HeritageAI: "
            "Could not go to previous slide."
        )

        print(
            "Error:",
            error
        )


# ============================================================
# GO TO SPECIFIC SLIDE
# ============================================================

def go_to_slide(number):

    try:

        number = int(number)

    except Exception:

        return

    try:

        if presentation is None:

            if not connect_powerpoint():

                return

        slide_count = (
            presentation.Slides.Count
        )

        if (
            number < 1
            or number > slide_count
        ):

            print()

            print(
                f"HeritageAI: "
                f"Slide {number} does not exist."
            )

            print(
                f"HeritageAI: "
                f"This presentation has "
                f"{slide_count} slides."
            )

            return

        if not slideshow_is_running():

            print(
                "HeritageAI: "
                "Slideshow is not running."
            )

            print(
                "HeritageAI: "
                "Starting presentation first..."
            )

            if not start_presentation():

                return

        view = (
            get_slideshow_view()
        )

        if view is None:

            return

        view.GotoSlide(number)

        print(
            f"HeritageAI: "
            f"Now on slide {number}."
        )

        focus_powerpoint()

    except Exception as error:

        print(
            "HeritageAI: "
            "Could not jump to slide."
        )

        print(
            "Error:",
            error
        )


# ============================================================
# STOP PRESENTATION
# ============================================================

def stop_presentation():

    if not slideshow_is_running():

        print(
            "HeritageAI: "
            "Slideshow is not running."
        )

        return

    try:

        view = (
            get_slideshow_view()
        )

        if view is not None:

            view.Exit()

        print(
            "HeritageAI: "
            "Presentation stopped."
        )

        focus_powerpoint()

    except Exception as error:

        print(
            "HeritageAI: "
            "Could not stop presentation."
        )

        print(
            "Error:",
            error
        )


# ============================================================
# EXTRACT SLIDE NUMBER
# ============================================================

def extract_slide_number(text):
    """Extract a slide number while tolerating common speech errors."""
    text = normalize_text(text)
    if not text:
        return None

    words = text.split()

    # Direct numeric form: "8", "slide 8", etc.
    for word in words:
        if word.isdigit():
            number = int(word)
            if 1 <= number <= 999:
                return number

    # Exact written numbers.
    for word in words:
        if word in NUMBER_WORDS:
            number = NUMBER_WORDS[word]
            if 1 <= number <= 999:
                return number

    # Contextual homophones produced by speech recognition.
    # These are only interpreted as numbers when the surrounding command
    # clearly looks like slide navigation.
    contextual_numbers = {
        "to": 2,
        "too": 2,
        "for": 4,
        "ate": 8,
        "won": 1,
        "wan": 1,
        "wun": 1,
    }

    navigation_context = any(phrase in text for phrase in (
        "slide", "number", "go to", "move to", "take me to",
        "take us to", "jump to", "show me", "bring me to",
        "go directly to", "open slide", "show slide",
    ))

    if navigation_context:
        for word in words:
            if word in contextual_numbers:
                return contextual_numbers[word]

    return None


# ============================================================
# PRESENTATION INTELLIGENCE
# ============================================================

TOPIC_STOP_WORDS = {
    "the", "a", "an", "to", "of", "on", "in", "at", "for",
    "me", "my", "please", "can", "could", "would", "you", "show",
    "find", "take", "bring", "go", "where", "did", "i", "we", "us",
    "discuss", "discussed", "talk", "talked", "about", "slide", "slides",
    "presentation", "presentations", "page", "pages", "is", "are", "was",
    "were", "this", "that", "these", "those", "from", "with", "and", "or"
}


def shape_text(shape):
    """Safely extract visible text from a PowerPoint shape."""
    try:
        if getattr(shape, "HasTextFrame", False):
            if shape.TextFrame.HasText:
                return str(shape.TextFrame.TextRange.Text or "")
    except Exception:
        pass

    try:
        if getattr(shape, "HasTable", False):
            values = []
            table = shape.Table
            for row in range(1, table.Rows.Count + 1):
                for col in range(1, table.Columns.Count + 1):
                    try:
                        values.append(str(table.Cell(row, col).Shape.TextFrame.TextRange.Text or ""))
                    except Exception:
                        pass
            return " ".join(values)
    except Exception:
        pass

    return ""


def extract_slide_text(slide):
    """Return all readable text from one slide."""
    parts = []
    try:
        for shape in slide.Shapes:
            text = shape_text(shape).strip()
            if text:
                parts.append(text)
    except Exception:
        pass
    return " ".join(parts)


def build_slide_index():
    """Read the selected presentation once and cache slide text."""
    global slide_index

    slide_index = []

    if presentation is None:
        return False

    try:
        count = presentation.Slides.Count
    except Exception:
        return False

    print()
    print("HeritageAI: Building presentation intelligence...")

    for number in range(1, count + 1):
        try:
            slide = presentation.Slides.Item(number)
            full_text = extract_slide_text(slide)
            normalized = normalize_text(full_text)

            # First meaningful text line is treated as a lightweight title.
            title = ""
            for line in full_text.splitlines():
                line = normalize_text(line)
                if line and len(line) >= 3:
                    title = line
                    break

            slide_index.append({
                "number": number,
                "title": title,
                "text": normalized,
            })
        except Exception:
            slide_index.append({
                "number": number,
                "title": "",
                "text": "",
            })

    print(
        "HeritageAI: Presentation intelligence ready for",
        len(slide_index),
        "slides."
    )
    return True


def topic_words(text):
    words = re.findall(r"[a-z0-9]+", normalize_text(text))
    return {w for w in words if w not in TOPIC_STOP_WORDS and len(w) >= 3}


def topic_command_candidate(command):
    """Decide whether a command is asking for a topic/subject."""
    command = normalize_text(command)
    if not command:
        return False

    topic_markers = [
        "find ",
        "find the ",
        "show ",
        "show me ",
        "take me to ",
        "bring me to ",
        "go to ",
        "jump to ",
        "where is ",
        "where did i discuss ",
        "where did we discuss ",
        "where did i talk about ",
        "where did we talk about ",
        "what slide is ",
        "show the slide about ",
        "go to the slide about ",
        "take me to the slide about ",
        "open the slide about ",
    ]

    return any(command.startswith(marker) for marker in topic_markers)


def extract_topic_query(command):
    """Remove navigation wording and return the subject being requested."""
    command = normalize_text(command)

    patterns = [
        r"^where did (?:i|we) (?:discuss|talk about) (.+)$",
        r"^where is (.+)$",
        r"^what slide is (.+)$",
        r"^find (?:the )?(.+)$",
        r"^show me (?:the )?(.+)$",
        r"^show (?:me )?(?:the )?slide (?:about|on) (.+)$",
        r"^show (?:the )?(?:slide )?(?:about|on) (.+)$",
        r"^go to (?:the )?(?:slide )?(?:about|on) (.+)$",
        r"^take me to (?:the )?(?:slide )?(?:about|on) (.+)$",
        r"^bring me to (?:the )?(?:slide )?(?:about|on) (.+)$",
        r"^jump to (?:the )?(?:slide )?(?:about|on) (.+)$",
        r"^open (?:the )?(?:slide )?(?:about|on) (.+)$",
        r"^show me (?:the )?(.+)$",
        r"^go to (?:the )?(.+)$",
        r"^take me to (?:the )?(.+)$",
        r"^bring me to (?:the )?(.+)$",
        r"^jump to (?:the )?(.+)$",
    ]

    for pattern in patterns:
        match = re.match(pattern, command)
        if match:
            query = match.group(1).strip()
            query = re.sub(r"^(slide|slides)\s+", "", query).strip()
            query = re.sub(r"\s+(slide|slides)$", "", query).strip()
            return query

    return command


def search_topic(command):
    """Find the best matching slide for a natural-language topic request."""
    global slide_index

    if not slide_index:
        build_slide_index()

    if not slide_index:
        return None

    query = extract_topic_query(command)
    query = normalize_text(query)
    query_tokens = topic_words(query)

    if not query or not query_tokens:
        return None

    ranked = []

    for item in slide_index:
        title = item["title"]
        text = item["text"]
        title_tokens = topic_words(title)
        text_tokens = topic_words(text)

        # Exact phrase is very strong.
        phrase_score = 1.0 if query in text else 0.0
        title_phrase_score = 1.0 if query in title else 0.0

        # Word overlap.
        title_overlap = (
            len(query_tokens & title_tokens) / len(query_tokens)
            if query_tokens else 0.0
        )
        text_overlap = (
            len(query_tokens & text_tokens) / len(query_tokens)
            if query_tokens else 0.0
        )

        # Fuzzy similarity helps with words like "methodology" vs a slightly
        # different phrase on the slide.
        title_fuzzy = SequenceMatcher(None, query, title).ratio() if title else 0.0
        text_fuzzy = SequenceMatcher(None, query, text[:1200]).ratio() if text else 0.0

        score = (
            title_phrase_score * 1.00
            + phrase_score * 0.80
            + title_overlap * 0.75
            + text_overlap * 0.45
            + title_fuzzy * 0.35
            + text_fuzzy * 0.10
        )

        ranked.append((score, item, title_overlap, text_overlap))

    ranked.sort(key=lambda x: x[0], reverse=True)

    if not ranked:
        return None

    best = ranked[0]
    second_score = ranked[1][0] if len(ranked) > 1 else 0.0

    # Require a meaningful match. Also require a reasonable gap when the
    # top two slides are close, so HeritageAI does not confidently jump to
    # a random slide.
    if best[0] < 0.55:
        return None

    if best[0] < 0.90 and (best[0] - second_score) < 0.08:
        return None

    return best[1]


def go_to_topic(command):
    """Search the presentation and jump to the best matching topic slide."""
    result = search_topic(command)

    if result is None:
        print()
        print("HeritageAI: I could not find a strong topic match.")
        return False

    number = result["number"]
    title = result["title"] or "matching content"

    print()
    print(
        f"HeritageAI: Topic match -> slide {number}"
    )
    print(
        "HeritageAI: Match ->",
        title
    )

    go_to_slide(number)
    return True


# ============================================================
# PROCESS COMMAND
# ============================================================

def _token_similarity(a, b):
    """Compare two words while tolerating small Vosk transcription errors."""
    return SequenceMatcher(None, a, b).ratio()


def _command_similarity(command, phrase):
    """Blend whole-command and token-level similarity."""
    command = normalize_text(command)
    phrase = normalize_text(phrase)

    if not command or not phrase:
        return 0.0

    if command == phrase:
        return 1.0

    whole = SequenceMatcher(None, command, phrase).ratio()

    command_tokens = command.split()
    phrase_tokens = phrase.split()

    matched = 0
    for token in command_tokens:
        if any(_token_similarity(token, target) >= 0.78 for target in phrase_tokens):
            matched += 1

    token_score = matched / max(len(command_tokens), len(phrase_tokens))
    return max(whole, token_score * 0.95)


def _best_command_match(command, phrases):
    best_phrase = None
    best_score = 0.0

    for phrase in phrases:
        score = _command_similarity(command, phrase)
        if score > best_score:
            best_score = score
            best_phrase = phrase

    return best_phrase, best_score


def _looks_like_next(command):
    command = normalize_text(command)
    _, score = _best_command_match(command, NEXT_PHRASES)

    # Short commands need a slightly stricter threshold than longer
    # natural-language commands.
    threshold = 0.78 if len(command.split()) >= 2 else 0.88
    return score >= threshold


def _looks_like_previous(command):
    command = normalize_text(command)
    _, score = _best_command_match(command, PREVIOUS_PHRASES)
    threshold = 0.78 if len(command.split()) >= 2 else 0.88
    return score >= threshold


def _looks_like_start(command):
    command = normalize_text(command)
    _, score = _best_command_match(command, START_PHRASES)
    threshold = 0.80 if len(command.split()) >= 2 else 0.90
    return score >= threshold


def _looks_like_stop(command):
    command = normalize_text(command)
    _, score = _best_command_match(command, STOP_PHRASES)
    threshold = 0.80 if len(command.split()) >= 2 else 0.90
    return score >= threshold


def _is_number_navigation(command):
    command = normalize_text(command)
    if not command:
        return False

    number = extract_slide_number(command)
    if number is None:
        return False

    # A bare number is useful while presenting: "five" means slide five.
    if command.isdigit() or command in NUMBER_WORDS:
        return True

    navigation_markers = (
        "slide", "number", "go to", "move to", "take me to",
        "take us to", "jump to", "show me", "show slide",
        "bring me to", "go directly to", "open slide",
    )

    return any(marker in command for marker in navigation_markers)


def process_command(text):
    """Interpret natural speech flexibly and execute one presentation action."""
    raw = normalize_text(text)

    if not raw:
        return True

    print()
    print("HeritageAI heard:", raw)

    command = correct_speech_words(raw)

    # Undo a few risky global substitutions if an older correction table
    # or recognizer has produced them. This keeps normal English intact.
    command = re.sub(r"\bgo two slide\b", "go to slide", command)
    command = re.sub(r"\bmove two slide\b", "move to slide", command)
    command = re.sub(r"\btake me two slide\b", "take me to slide", command)
    command = re.sub(r"\bjump two slide\b", "jump to slide", command)
    command = re.sub(r"\bbring me two slide\b", "bring me to slide", command)
    command = normalize_text(command)

    print("HeritageAI interpreted:", command)

    # --------------------------------------------------------
    # 1. STOP / END
    # --------------------------------------------------------
    if command in STOP_PHRASES or _looks_like_stop(command):
        print("HeritageAI: Stop command detected.")
        stop_presentation()
        return False

    # --------------------------------------------------------
    # 2. START / BEGIN
    # --------------------------------------------------------
    if command in START_PHRASES or _looks_like_start(command):
        print("HeritageAI: Start command detected.")
        start_presentation()
        return True

    # --------------------------------------------------------
    # 3. DIRECT SLIDE NUMBER
    # --------------------------------------------------------
    if _is_number_navigation(command):
        number = extract_slide_number(command)
        print(f"HeritageAI: Going directly to slide {number}...")
        go_to_slide(number)
        return True

    # --------------------------------------------------------
    # 4. NEXT / FORWARD
    # --------------------------------------------------------
    if command in NEXT_PHRASES or _looks_like_next(command):
        print("HeritageAI: Next-slide command detected.")
        next_slide()
        return True

    # --------------------------------------------------------
    # 5. PREVIOUS / BACK
    # --------------------------------------------------------
    if command in PREVIOUS_PHRASES or _looks_like_previous(command):
        print("HeritageAI: Previous-slide command detected.")
        previous_slide()
        return True

    # --------------------------------------------------------
    # 6. TOPIC / PRESENTATION INTELLIGENCE
    # --------------------------------------------------------
    if topic_command_candidate(command):
        print("HeritageAI: Topic search detected.")
        go_to_topic(command)
        return True

    # --------------------------------------------------------
    # 7. Natural topic phrasing without an exact marker.
    # --------------------------------------------------------
    topic_words_in_command = topic_words(command)
    topic_like_markers = (
        "methodology", "method", "results", "conclusion",
        "architecture", "introduction", "problem", "solution",
        "objective", "objectives", "background", "discussion",
        "recommendation", "recommendations", "summary",
    )

    if topic_words_in_command and any(
        marker in command for marker in topic_like_markers
    ):
        print("HeritageAI: Natural topic request detected.")
        go_to_topic("show " + command)
        return True

    print("HeritageAI: Command not recognized.")
    return True


# ============================================================
# AUDIO CALLBACK
# ============================================================

def audio_callback(
    indata,
    frames,
    time_info,
    status
):

    try:

        audio_queue.put(
            bytes(indata)
        )

    except Exception:

        pass


# ============================================================
# LOAD VOSK MODEL
# ============================================================

def load_voice_model():

    global model

    print()
    print(
        "HeritageAI: "
        "Loading voice model..."
    )

    possible_paths = [
        MODEL_PATH_INTERNAL,
        MODEL_PATH_PACKAGED,
    ]

    model_path = None

    for path in possible_paths:

        print(
            "Checking model:",
            path
        )

        if os.path.isdir(path):

            model_path = path

            break

    if model_path is None:

        print()
        print(
            "HeritageAI: "
            "Voice model folder was not found."
        )

        print()
        print(
            "Expected one of:"
        )

        for path in possible_paths:

            print(
                path
            )

        return False

    try:

        model = vosk.Model(
            model_path
        )

        print()
        print(
            "HeritageAI: "
            "Voice model loaded."
        )

        return True

    except Exception as error:

        print()
        print(
            "HeritageAI: "
            "Could not load voice model."
        )

        print(
            "Error:",
            error
        )

        return False


# ============================================================
# CREATE RECOGNIZER
# ============================================================

def create_recognizer():

    global recognizer

    try:

        # IMPORTANT:
        #
        # The old version used a restrictive grammar.
        #
        # That made recognition more predictable,
        # but it prevented natural variations.
        #
        # We now allow normal Vosk recognition and
        # interpret the result ourselves in process_command().

        recognizer = (
            vosk.KaldiRecognizer(
                model,
                sample_rate
            )
        )

        recognizer.SetWords(
            True
        )

        return True

    except Exception as error:

        print()
        print(
            "HeritageAI: "
            "Could not create speech recognizer."
        )

        print(
            "Error:",
            error
        )

        return False


# ============================================================
# MAIN VOICE ENGINE
# ============================================================

def main():

    print()
    print(
        "======================================"
    )

    print(
        "             HERITAGE AI"
    )

    print(
        "     Voice Presentation Assistant"
    )

    print(
        "======================================"
    )

    # --------------------------------------------------------
    # LOAD MODEL
    # --------------------------------------------------------

    if not load_voice_model():

        input(
            "Press ENTER to close..."
        )

        return

    # --------------------------------------------------------
    # MICROPHONE
    # --------------------------------------------------------

    if not find_microphone():

        input(
            "Press ENTER to close..."
        )

        return

    # --------------------------------------------------------
    # POWERPOINT
    # --------------------------------------------------------

    if presentation is None:

        if not connect_powerpoint():

            return

    # --------------------------------------------------------
    # RECOGNIZER
    # --------------------------------------------------------

    if not create_recognizer():

        input(
            "Press ENTER to close..."
        )

        return

    # --------------------------------------------------------
    # COMMANDS
    # --------------------------------------------------------

    print()
    print(
        "Commands:"
    )

    print(
        "  Start presentation"
    )

    print(
        "  Go to slide 3"
    )

    print(
        "  Go to slide three"
    )

    print(
        "  Next"
    )

    print(
        "  Go forward"
    )

    print(
        "  Previous"
    )

    print(
        "  Go back"
    )

    print(
        "  Show the methodology"
    )

    print(
        "  Take me to the conclusion"
    )

    print(
        "  Where did I discuss networking"
    )

    print(
        "  Stop"
    )

    print()
    print(
        "HeritageAI is ready."
    )

    # --------------------------------------------------------
    # MICROPHONE
    # --------------------------------------------------------

    print()
    print(
        "HeritageAI: "
        "Testing microphone..."
    )

    try:

        with sd.RawInputStream(
            samplerate=sample_rate,
            blocksize=8000,
            device=microphone_index,
            dtype="int16",
            channels=1,
            callback=audio_callback
        ):

            print(
                "HeritageAI: "
                "Microphone opened successfully."
            )

            print()
            print(
                "Listening..."
            )

            print()

            running = True

            while running:

                try:

                    data = (
                        audio_queue.get(
                            timeout=1
                        )
                    )

                except queue.Empty:

                    continue

                try:

                    if recognizer.AcceptWaveform(
                        data
                    ):

                        result = json.loads(
                            recognizer.Result()
                        )

                        text = result.get(
                            "text",
                            ""
                        )

                        if text:

                            running = (
                                process_command(
                                    text
                                )
                            )

                except Exception as error:

                    print()
                    print(
                        "HeritageAI: "
                        "Speech processing error:"
                    )

                    print(
                        "Error:",
                        error
                    )

    except KeyboardInterrupt:

        print()

    except Exception as error:

        print()
        print(
            "HeritageAI: "
            "Could not open microphone."
        )

        print(
            "Error:",
            error
        )

    print()
    print(
        "HeritageAI stopped."
    )


# ============================================================
# USER INTERFACE
# ============================================================

def launch_interface():
    import tkinter as tk
    from tkinter import filedialog, messagebox
    import os

    BG = "#080808"
    PANEL = "#111111"
    PANEL_2 = "#151515"
    GOLD = "#D4AF37"
    GOLD_LIGHT = "#E5C65A"
    WHITE = "#F4F4F4"
    MUTED = "#929292"
    BORDER = "#292929"
    GREEN = "#63B77A"

    selected_file = {"path": None}

    def start_selected_presentation():
        global powerpoint, presentation

        file_path = selected_file["path"]

        if not file_path:
            messagebox.showwarning(
                "No Presentation Selected",
                "Please choose a PowerPoint presentation first."
            )
            return

        try:
            status_label.config(text="OPENING PRESENTATION...", fg=GOLD)
            status_dot.config(fg=GOLD)
            root.update_idletasks()

            file_path = os.path.abspath(os.path.normpath(file_path))

            if not os.path.isfile(file_path):
                raise FileNotFoundError(file_path)

            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            try:
                powerpoint.Visible = True
            except Exception:
                pass

            presentation = powerpoint.Presentations.Open(
                FileName=file_path,
                ReadOnly=False,
                Untitled=False,
                WithWindow=True
            )

            time.sleep(1)

            build_slide_index()
            root.destroy()
            main()

        except Exception as error:
            status_label.config(
                text="COULD NOT OPEN PRESENTATION",
                fg="#D96B6B"
            )
            status_dot.config(fg="#D96B6B")
            messagebox.showerror(
                "HeritageAI",
                "Could not open the presentation.\n\n" + str(error)
            )

    def choose_presentation():
        file_path = filedialog.askopenfilename(
            title="Choose a PowerPoint Presentation",
            filetypes=[
                ("PowerPoint Presentation", "*.pptx"),
                ("PowerPoint 97-2003", "*.ppt"),
                ("All Files", "*.*")
            ]
        )

        if not file_path:
            return

        selected_file["path"] = file_path
        filename = os.path.basename(file_path)

        file_name_label.config(text=filename, fg=WHITE)
        file_path_label.config(text=file_path, fg=MUTED)
        status_label.config(text="PRESENTATION READY", fg=GREEN)
        status_dot.config(fg=GREEN)

        start_button.config(
            state="normal",
            bg=GOLD,
            fg="#080808",
            activebackground=GOLD_LIGHT,
            activeforeground="#080808",
            cursor="hand2"
        )

    root = tk.Tk()
    root.title("HeritageAI")
    root.geometry("820x560")
    root.minsize(760, 520)
    root.configure(bg=BG)

    root.update_idletasks()
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - 820) // 2
    y = (screen_height - 560) // 2
    root.geometry(f"820x560+{x}+{y}")

    top_bar = tk.Frame(root, bg=BG, height=75)
    top_bar.pack(fill="x", padx=42, pady=(28, 0))

    brand_frame = tk.Frame(top_bar, bg=BG)
    brand_frame.pack(side="left", anchor="w")

    tk.Label(
        brand_frame,
        text="HERITAGE",
        font=("Segoe UI", 18, "bold"),
        fg=WHITE,
        bg=BG
    ).pack(side="left")

    tk.Label(
        brand_frame,
        text="AI",
        font=("Segoe UI", 18, "bold"),
        fg=GOLD,
        bg=BG
    ).pack(side="left", padx=(3, 0))

    tk.Label(
        brand_frame,
        text="  POWERPOINT VOICE CONTROL",
        font=("Segoe UI", 8, "bold"),
        fg=MUTED,
        bg=BG
    ).pack(side="left", padx=(12, 0), pady=(5, 0))

    content = tk.Frame(root, bg=BG)
    content.pack(fill="both", expand=True, padx=42, pady=(10, 0))

    tk.Frame(content, bg=GOLD, height=2).pack(fill="x", pady=(5, 35))

    tk.Label(
        content,
        text="Present With Your Voice.",
        font=("Segoe UI", 30, "bold"),
        fg=WHITE,
        bg=BG
    ).pack(anchor="w")

    tk.Label(
        content,
        text="Control your PowerPoint presentation without leaving your flow.",
        font=("Segoe UI", 11),
        fg=MUTED,
        bg=BG
    ).pack(anchor="w", pady=(8, 28))

    presentation_panel = tk.Frame(
        content,
        bg=PANEL,
        highlightbackground=BORDER,
        highlightthickness=1
    )
    presentation_panel.pack(fill="x")

    tk.Label(
        presentation_panel,
        text="PRESENTATION",
        font=("Segoe UI", 9, "bold"),
        fg=GOLD,
        bg=PANEL
    ).pack(anchor="w", padx=26, pady=(23, 10))

    file_area = tk.Frame(
        presentation_panel,
        bg=PANEL_2,
        highlightbackground="#242424",
        highlightthickness=1
    )
    file_area.pack(fill="x", padx=26, pady=(0, 20))

    tk.Frame(file_area, bg=GOLD, width=4).pack(side="left", fill="y")

    file_text_area = tk.Frame(file_area, bg=PANEL_2)
    file_text_area.pack(side="left", fill="both", expand=True, padx=18, pady=15)

    file_name_label = tk.Label(
        file_text_area,
        text="No presentation selected",
        font=("Segoe UI", 11, "bold"),
        fg=MUTED,
        bg=PANEL_2,
        anchor="w"
    )
    file_name_label.pack(fill="x")

    file_path_label = tk.Label(
        file_text_area,
        text="Choose a .pptx or .ppt file from your computer",
        font=("Segoe UI", 8),
        fg="#6F6F6F",
        bg=PANEL_2,
        anchor="w"
    )
    file_path_label.pack(fill="x", pady=(4, 0))

    tk.Button(
        presentation_panel,
        text="Choose Presentation",
        font=("Segoe UI", 10, "bold"),
        bg=PANEL,
        fg=WHITE,
        activebackground=PANEL_2,
        activeforeground=GOLD,
        relief="flat",
        bd=0,
        cursor="hand2",
        padx=20,
        pady=10,
        command=choose_presentation
    ).pack(anchor="w", padx=26, pady=(0, 23))

    lower = tk.Frame(content, bg=BG)
    lower.pack(fill="x", pady=(22, 0))

    status_frame = tk.Frame(lower, bg=BG)
    status_frame.pack(side="left", anchor="w")

    status_dot = tk.Label(
        status_frame,
        text="●",
        font=("Segoe UI", 9),
        fg=MUTED,
        bg=BG
    )
    status_dot.pack(side="left")

    status_label = tk.Label(
        status_frame,
        text="SELECT A PRESENTATION TO BEGIN",
        font=("Segoe UI", 8, "bold"),
        fg=MUTED,
        bg=BG
    )
    status_label.pack(side="left", padx=(7, 0))

    start_button = tk.Button(
        lower,
        text="Start Presentation  →",
        font=("Segoe UI", 10, "bold"),
        bg="#292929",
        fg="#777777",
        activebackground="#292929",
        activeforeground="#777777",
        relief="flat",
        bd=0,
        cursor="arrow",
        padx=24,
        pady=12,
        state="disabled",
        command=start_selected_presentation
    )
    start_button.pack(side="right")

    commands_panel = tk.Frame(content, bg=BG)
    commands_panel.pack(fill="x", pady=(25, 0))

    tk.Label(
        commands_panel,
        text="VOICE CONTROLS",
        font=("Segoe UI", 8, "bold"),
        fg=GOLD,
        bg=BG
    ).pack(anchor="w")

    tk.Label(
        commands_panel,
        text="Next   •   Previous   •   Go to slide   •   Find a topic   •   Start   •   Stop",
        font=("Segoe UI", 9),
        fg="#777777",
        bg=BG
    ).pack(anchor="w", pady=(6, 0))

    footer = tk.Frame(root, bg=BG)
    footer.pack(fill="x", padx=42, pady=(0, 24))

    tk.Frame(footer, bg="#202020", height=1).pack(fill="x", pady=(0, 12))

    tk.Label(
        footer,
        text="HERITAGEAI  •  PRESENTATION, REIMAGINED",
        font=("Segoe UI", 7, "bold"),
        fg="#555555",
        bg=BG
    ).pack(anchor="w")

    root.bind(
        "<Return>",
        lambda event: start_selected_presentation() if selected_file["path"] else None
    )

    root.bind("<Escape>", lambda event: root.destroy())

    root.mainloop()


# ============================================================
# START PROGRAM
# ============================================================

if __name__ == "__main__":

    launch_interface()

