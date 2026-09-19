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


if getattr(sys, "frozen", False):

    MODEL_BASE_PATH = sys._MEIPASS

else:

    MODEL_BASE_PATH = BASE_PATH


MODEL_PATH_INTERNAL = os.path.join(
    MODEL_BASE_PATH,
    "_internal",
    "vosk-model-small-en-us-0.15"
)

MODEL_PATH_LOCAL = os.path.join(
    MODEL_BASE_PATH,
    "vosk-model-small-en-us-0.15"
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


# ============================================================
# NUMBER WORDS
# ============================================================

NUMBER_WORDS = {

    "zero": 0,

    "one": 1,
    "two": 2,
    "three": 3,
    "four": 4,
    "five": 5,
    "six": 6,
    "seven": 7,
    "eight": 8,
    "nine": 9,
    "ten": 10,

    "eleven": 11,
    "twelve": 12,
    "thirteen": 13,
    "fourteen": 14,
    "fifteen": 15,
    "sixteen": 16,
    "seventeen": 17,
    "eighteen": 18,
    "nineteen": 19,
    "twenty": 20,

    "twenty one": 21,
    "twenty two": 22,
    "twenty three": 23,
    "twenty four": 24,
    "twenty five": 25,
    "twenty six": 26,
    "twenty seven": 27,
    "twenty eight": 28,
    "twenty nine": 29,

    "thirty": 30,
    "thirty one": 31,
    "thirty two": 32,
    "thirty three": 33,
    "thirty four": 34,
    "thirty five": 35,
    "thirty six": 36,
    "thirty seven": 37,
    "thirty eight": 38,
    "thirty nine": 39,

    "forty": 40,
    "forty one": 41,
    "forty two": 42,
    "forty three": 43,
    "forty four": 44,
    "forty five": 45,
    "forty six": 46,
    "forty seven": 47,
    "forty eight": 48,
    "forty nine": 49,

    "fifty": 50,
    "fifty one": 51,
    "fifty two": 52,
    "fifty three": 53,
    "fifty four": 54,
    "fifty five": 55,
    "fifty six": 56,
    "fifty seven": 57,
    "fifty eight": 58,
    "fifty nine": 59,

    "sixty": 60,
    "sixty one": 61,
    "sixty two": 62,
    "sixty three": 63,
    "sixty four": 64,
    "sixty five": 65,
    "sixty six": 66,
    "sixty seven": 67,
    "sixty eight": 68,
    "sixty nine": 69,

    "seventy": 70,
    "seventy one": 71,
    "seventy two": 72,
    "seventy three": 73,
    "seventy four": 74,
    "seventy five": 75,
    "seventy six": 76,
    "seventy seven": 77,
    "seventy eight": 78,
    "seventy nine": 79,

    "eighty": 80,
    "eighty one": 81,
    "eighty two": 82,
    "eighty three": 83,
    "eighty four": 84,
    "eighty five": 85,
    "eighty six": 86,
    "eighty seven": 87,
    "eighty eight": 88,
    "eighty nine": 89,

    "ninety": 90,
    "ninety one": 91,
    "ninety two": 92,
    "ninety three": 93,
    "ninety four": 94,
    "ninety five": 95,
    "ninety six": 96,
    "ninety seven": 97,
    "ninety eight": 98,
    "ninety nine": 99,
}


# ============================================================
# SPEECH CORRECTIONS
# ============================================================

SPEECH_CORRECTIONS = {

    # -------------------------
    # Heritage
    # -------------------------

    "herit": "heritage",
    "heritagee": "heritage",
    "heritages": "heritage",
    "her itage": "heritage",
    "hairitage": "heritage",
    "heritage ai": "heritage",
    "heritagea": "heritage",
    "heritage i": "heritage",

    # -------------------------
    # Next
    # -------------------------

    "neck": "next",
    "necks": "next",
    "nest": "next",
    "nests": "next",
    "neste": "next",
    "nexst": "next",
    "nest": "next",
    "nests": "next",
    "neste": "next",
    "nex": "next",
    "nextt": "next",
    "nexted": "next",
    "nexx": "next",

    # -------------------------
    # Previous
    # -------------------------

    "previews": "previous",
    "preview": "previous",
    "previously": "previous",
    "previse": "previous",
    "previus": "previous",
    "prev": "previous",

    # -------------------------
    # Start
    # -------------------------

    "stat": "start",
    "stark": "start",
    "starts": "start",
    "started": "start",
    "starr": "start",

    # -------------------------
    # Stop
    # -------------------------

    "stops": "stop",
    "stopped": "stop",
    "shop": "stop",
    "stap": "stop",
    "stopp": "stop",

    # -------------------------
    # Backward
    # -------------------------

    "backwards": "backward",
    "backword": "backward",

    # -------------------------
    # Presentation
    # -------------------------

    "presentations": "presentation",
    "presenting": "presentation",
    "presented": "presentation",

    # -------------------------
    # Slide
    # -------------------------

    "slides": "slide",
    "slight": "slide",
    "slights": "slide",
    "slyde": "slide",

    # -------------------------
    # Go
    # -------------------------

    "goo": "go",
    "goh": "go",

    # -------------------------
    # Number
    # -------------------------

    "numbers": "number",
    "numb": "number",
}


# ============================================================
# SMART SPEECH VARIANTS
#
# Common words that Vosk may confuse with command words.
# These are applied carefully so ordinary speech is not
# accidentally turned into a command.
# ============================================================

NUMBER_SPEECH_VARIANTS = {
    # ONE
    "won": "one",
    "wan": "one",
    "wun": "one",

    # TWO
    "too": "two",
    "tu": "two",
    "tue": "two",
    "to": "two",

    # THREE
    "tree": "three",
    "threes": "three",
    "thre": "three",
    "free": "three",
    "feree": "three",

    # FOUR
    "for": "four",
    "fo": "four",
    "fore": "four",
    "phor": "four",

    # EIGHT
    "ate": "eight",
    "ait": "eight",
    "aight": "eight",
    "eigt": "eight",
    "eigth": "eight",
    "eightt": "eight",

    # NINE
    "nein": "nine",
    "nien": "nine",
    "nyn": "nine",
    "nain": "nine",
    "night": "nine",

    # TEN
    "then": "ten",
    "tin": "ten",
    "tenn": "ten",
    "tend": "ten",

    # ELEVEN
    "elevn": "eleven",
    "elevent": "eleven",
    "elebin": "eleven",
    "eleben": "eleven",
    "elevenn": "eleven",

    # TWELVE
    "twelv": "twelve",
    "twelf": "twelve",
    "twelvth": "twelve",
    "twelbe": "twelve",
    "twelb": "twelve",

    # OTHER COMMON NUMBER ERRORS
    "fiv": "five",
    "fife": "five",
    "twentee": "twenty",
    "thirtee": "thirty",
    "fourty": "forty",
    "fivty": "fifty",
    "sixty": "sixty",
    "seventy": "seventy",
    "eighty": "eighty",
    "ninty": "ninety",
}

COMMAND_FILLER_WORDS = {
    "please", "can", "could", "would", "will", "you",
    "kindly", "just", "now", "me", "the", "a", "an",
    "to", "my", "for", "us", "let", "lets", "i",
}


# ============================================================
# FUZZY NUMBER CORRECTION
# ============================================================

def correct_number_words(words):

    # Canonical number vocabulary for every integer from 0 to 99.
    # We use fuzzy matching on individual number words so Vosk can
    # recover from natural pronunciation/transcription variations.
    single_number_words = {
        "zero", "one", "two", "three", "four", "five",
        "six", "seven", "eight", "nine", "ten", "eleven",
        "twelve", "thirteen", "fourteen", "fifteen",
        "sixteen", "seventeen", "eighteen", "nineteen",
        "twenty", "thirty", "forty", "fifty", "sixty",
        "seventy", "eighty", "ninety"
    }

    corrected = []

    for word in words:

        if word in single_number_words:
            corrected.append(word)
            continue

        if word in NUMBER_SPEECH_VARIANTS:
            corrected.append(NUMBER_SPEECH_VARIANTS[word])
            continue

        best_word = None
        best_score = 0.0

        for number_word in single_number_words:
            score = similarity(word, number_word)

            if score > best_score:
                best_score = score
                best_word = number_word

        # Conservative threshold: only words that are reasonably
        # close to a real number word are corrected.
        if best_word is not None and best_score >= 0.76:
            corrected.append(best_word)
        else:
            corrected.append(word)

    return corrected


# ============================================================
# NUMBER EXTRACTION HELPERS
# ============================================================

def number_from_words(number_tokens):

    if not number_tokens:
        return None

    # Correct each spoken number component. This handles things like
    # "twentee tree" -> "twenty three" and "thirtee for" -> "thirty four".
    corrected = correct_number_words(number_tokens)

    # Only accept tokens that are genuine canonical number words.
    if any(token not in NUMBER_WORDS for token in corrected):
        return None

    phrase = " ".join(corrected)

    if phrase in NUMBER_WORDS:
        number = NUMBER_WORDS[phrase]
        return number if 1 <= number <= 99 else None

    # Also allow a tens word + a unit word even if the exact phrase
    # was not explicitly present in the dictionary.
    tens = {
        "twenty": 20, "thirty": 30, "forty": 40,
        "fifty": 50, "sixty": 60, "seventy": 70,
        "eighty": 80, "ninety": 90
    }

    units = {
        "one": 1, "two": 2, "three": 3, "four": 4,
        "five": 5, "six": 6, "seven": 7, "eight": 8,
        "nine": 9
    }

    if len(corrected) == 2 and corrected[0] in tens and corrected[1] in units:
        return tens[corrected[0]] + units[corrected[1]]

    return None


# ============================================================
# SMART COMMAND TEXT
# ============================================================

def smart_command_text(text):

    text = normalize_text(text)

    if not text:
        return ""

    words = text.split()

    # Only correct ambiguous number words when they are likely
    # being used as a slide number. This avoids changing
    # normal phrases such as "go to" into "go two".
    number_context = any(
        phrase in text
        for phrase in (
            "number",
            "slide",
            "go to",
            "goto",
            "jump to",
            "move to",
            "take me to",
            "bring me to",
            "show",
            "open",
        )
    )

    if number_context:
        words = correct_number_words(words)

    text = " ".join(words)

    # Strip conversational filler only for matching.
    # This lets phrases such as "can you please start" work.
    words = text.split()
    filtered = [
        word
        for word in words
        if word not in COMMAND_FILLER_WORDS
    ]

    if filtered:
        return " ".join(filtered)

    return text


# ============================================================
# COMMAND PHRASES
# ============================================================

NEXT_PHRASES = [

    "next",
    "go next",
    "next slide",
    "go to next",
    "go to next slide",
    "move next",
    "move to next",
    "move to next slide",
    "move forward",
    "go forward",
    "forward",
    "continue",
    "proceed",
    "show next",
    "show the next",
    "show the next slide",
    "take me to next",
    "take me to the next",
    "take me to the next slide",
    "bring up the next",
    "bring up the next slide",
]


PREVIOUS_PHRASES = [

    "previous",
    "previous slide",
    "go previous",
    "go to previous",
    "go to previous slide",
    "move previous",
    "move to previous",
    "move to previous slide",
    "go back",
    "back",
    "move back",
    "move backward",
    "backward",
    "return",
    "show previous",
    "show the previous",
    "show the previous slide",
    "take me to previous",
    "take me to the previous",
    "take me to the previous slide",
]


START_PHRASES = [

    "start",
    "start presentation",
    "start the presentation",
    "begin",
    "begin presentation",
    "begin the presentation",
    "present",
    "presentation",
    "launch presentation",
    "launch the presentation",
    "open presentation",
    "open the presentation",
    "commence",
    "commence presentation",
    "commence the presentation",
    "let us start",
    "lets start",
    "let us begin",
    "lets begin",
]


STOP_PHRASES = [

    "stop",
    "stop presentation",
    "stop the presentation",
    "end",
    "end presentation",
    "end the presentation",
    "finish",
    "finish presentation",
    "finish the presentation",
    "exit",
    "exit presentation",
    "exit the presentation",
    "close presentation",
    "close the presentation",
    "quit presentation",
    "quit the presentation",
]


# ============================================================
# NORMALIZE TEXT
# ============================================================

def normalize_text(text):

    if not text:
        return ""

    text = text.lower().strip()

    text = text.replace(
        "’",
        "'"
    )

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
# CORRECT COMMON SPEECH ERRORS
# ============================================================

def correct_speech_words(text):

    text = normalize_text(text)

    if not text:
        return ""

    # Multi-word corrections first.
    for wrong, correct in sorted(
        SPEECH_CORRECTIONS.items(),
        key=lambda item: len(item[0]),
        reverse=True
    ):

        if " " in wrong:

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
# SIMILARITY
# ============================================================

def similarity(a, b):

    a = normalize_text(a)
    b = normalize_text(b)

    if not a or not b:
        return 0.0

    return SequenceMatcher(
        None,
        a,
        b
    ).ratio()


# ============================================================
# EXACT COMMAND MATCHING
# ============================================================

def exact_phrase_match(
    text,
    phrases
):

    text = normalize_text(text)

    if not text:
        return False

    # Exact complete command.
    if text in phrases:
        return True

    # Exact phrase inside a longer command.
    padded_text = f" {text} "

    for phrase in phrases:

        padded_phrase = f" {phrase} "

        if padded_phrase in padded_text:

            return True

    return False


# ============================================================
# SAFE FUZZY COMMAND MATCHING
# ============================================================

def safe_phrase_match(
    text,
    phrases,
    threshold=0.82
):

    text = normalize_text(text)

    if not text:
        return False

    # Always prefer exact matching.
    if exact_phrase_match(
        text,
        phrases
    ):

        return True

    words = text.split()

    # --------------------------------------------------------
    # IMPORTANT SAFETY RULE
    # --------------------------------------------------------
    # Do NOT fuzzy-match a one-word command against another
    # one-word command.
    #
    # This prevents:
    #
    # next -> stop
    # start -> stop
    # back -> next
    #
    # etc.
    # --------------------------------------------------------

    if len(words) == 1:

        # Only allow fuzzy matching against phrases
        # that are also one word.
        for phrase in phrases:

            if " " in phrase:
                continue

            if len(phrase) < 4:
                continue

            if similarity(
                text,
                phrase
            ) >= threshold:

                return True

        return False

    # For longer phrases, compare the complete phrase.
    best_score = 0.0

    for phrase in phrases:

        score = similarity(
            text,
            phrase
        )

        if score > best_score:

            best_score = score

    return best_score >= threshold


# ============================================================
# MICROPHONE
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

            if (
                device[
                    "max_input_channels"
                ] > 0
            ):

                microphone_index = (
                    preferred_index
                )

        if microphone_index is None:

            for i, device in enumerate(
                devices
            ):

                if (
                    device[
                        "max_input_channels"
                    ] > 0
                ):

                    microphone_index = i

                    break

        if microphone_index is None:

            print()
            print(
                "HeritageAI: No microphone was found."
            )

            return False

        print()
        print(
            "HeritageAI: Using microphone:",
            microphone_index
        )

        sample_rate = 44100

        print(
            "HeritageAI: Using sample rate:",
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
# FIND POWERPOINT
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

        def enum_windows(hwnd, extra):

            try:

                title = (
                    win32gui.GetWindowText(
                        hwnd
                    )
                )

                if not title:
                    return

                title_lower = title.lower()

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

        # Otherwise normal PowerPoint.
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

        try:

            active = (
                powerpoint.ActivePresentation
            )

            if active is not None:

                return active

        except Exception:

            pass

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
# CONNECT POWERPOINT
# ============================================================

def connect_powerpoint():

    global powerpoint
    global presentation

    print()
    print(
        "HeritageAI: Connecting to PowerPoint..."
    )

    try:

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

            try:
                powerpoint.Visible = True
            except Exception:
                pass

            print(
                "HeritageAI: "
                "PowerPoint application started."
            )

        presentation = (
            get_active_presentation()
        )

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
                "        ",
                presentation_name
            )

        else:

            if not os.path.exists(
                PRESENTATION_FILE
            ):

                print()
                print(
                    "HeritageAI: "
                    "No PowerPoint presentation is open."
                )

                print(
                    "HeritageAI also could not find:"
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
                    ReadOnly=False,
                    Untitled=False,
                    WithWindow=True
                )
            )

            time.sleep(1)

        if presentation is None:

            print(
                "HeritageAI: "
                "Presentation unavailable."
            )

            return False

        try:

            name = presentation.Name

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
            "HeritageAI: Using presentation:"
        )

        print(
            "        ",
            name
        )

        print(
            "HeritageAI: Number of slides:",
            slide_count
        )

        print()
        print(
            "HeritageAI: "
            "Connected to PowerPoint."
        )

        focus_powerpoint()

        return True

    except Exception as error:

        print()
        print(
            "HeritageAI: "
            "Could not connect to PowerPoint."
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
        "HeritageAI: Starting presentation..."
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

        try:
            powerpoint.Visible = True
        except Exception:
            pass

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

        view = get_slideshow_view()

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

        view = get_slideshow_view()

        if view is None:
            return

        view.Next()

        print(
            "HeritageAI: Next slide."
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

        view = get_slideshow_view()

        if view is None:
            return

        view.Previous()

        print(
            "HeritageAI: Previous slide."
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

        view = get_slideshow_view()

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

        view = get_slideshow_view()

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

    text = normalize_text(text)

    if not text:
        return None

    # First accept an explicitly spoken/written numeric value.
    for token in text.split():
        if token.isdigit():
            number = int(token)
            if 1 <= number <= 99:
                return number

    # Remove common navigation words. This is important because words
    # like "to" must NOT be interpreted as the number "two".
    tokens = text.split()

    navigation_words = {
        "number", "slide", "slides", "go", "goto", "jump",
        "move", "take", "bring", "show", "open", "to",
        "the", "a", "an", "please", "can", "could", "would",
        "you", "kindly", "me", "my", "now", "just", "on",
        "into", "onto", "presentation"
    }

    number_tokens = [
        token for token in tokens
        if token not in navigation_words
    ]

    # Try the whole remaining phrase first.
    number = number_from_words(number_tokens)
    if number is not None:
        return number

    # If Vosk inserted an extra word, try each contiguous 1-2 word
    # section. This keeps recognition tolerant without accepting a
    # completely unrelated sentence as a slide number.
    for size in (2, 1):
        for i in range(len(number_tokens) - size + 1):
            candidate = number_tokens[i:i + size]
            number = number_from_words(candidate)
            if number is not None:
                return number

    return None


# ============================================================
# CHECK SLIDE NUMBER COMMAND
# ============================================================

def is_slide_number_command(text):

    text = normalize_text(text)

    if not text:
        return False

    number = extract_slide_number(text)

    if number is None:
        return False

    words = text.split()

    # A bare number/number word is a direct slide command.
    if len(words) == 1:
        if words[0].isdigit():
            return 1 <= int(words[0]) <= 99

        if words[0] in NUMBER_WORDS:
            return True

        # Also accept a single fuzzy-spoken number.
        return number_from_words(words) is not None

    # Strong navigation indicators.
    indicators = {
        "number", "slide", "slides", "go to", "goto",
        "jump to", "move to", "take me to", "bring me to",
        "show", "open", "take me", "bring me"
    }

    if any(indicator in text for indicator in indicators):
        return True

    # Short commands such as "twenty three" are also accepted.
    if len(words) <= 2 and number_from_words(words) is not None:
        return True

    return False


# ============================================================
# PROCESS COMMAND
# ============================================================

def process_command(text):

    original_text = text

    text = correct_speech_words(text)

    if not text:
        return True

    print()
    print(
        "HeritageAI heard:",
        original_text
    )

    command = smart_command_text(text)

    if not command:
        return True

    print(
        "HeritageAI command:",
        command
    )

    # ========================================================
    # 1. SLIDE NUMBER
    # ========================================================

    if is_slide_number_command(command):

        number = extract_slide_number(command)

        if number is not None:

            print()
            print(
                f"HeritageAI: Going directly to slide {number}..."
            )

            go_to_slide(number)
            return True

    # ========================================================
    # 2. NEXT
    # ========================================================

    if exact_phrase_match(command, NEXT_PHRASES):
        next_slide()
        return True

    if safe_phrase_match(command, NEXT_PHRASES, threshold=0.76):
        next_slide()
        return True

    # Natural-language intent: anything clearly asking for
    # the next/forward slide should work.
    next_intent_words = {
        "next", "forward", "continue", "proceed",
        "advance", "following"
    }

    if any(word in next_intent_words for word in command.split()):
        if not any(word in {"previous", "back", "backward", "stop"}
                   for word in command.split()):
            next_slide()
            return True

    # ========================================================
    # 3. PREVIOUS
    # ========================================================

    if exact_phrase_match(command, PREVIOUS_PHRASES):
        previous_slide()
        return True

    if safe_phrase_match(command, PREVIOUS_PHRASES, threshold=0.76):
        previous_slide()
        return True

    previous_intent_words = {
        "previous", "back", "backward", "return", "behind"
    }

    if any(word in previous_intent_words for word in command.split()):
        if not any(word in {"next", "forward", "stop"}
                   for word in command.split()):
            previous_slide()
            return True

    # ========================================================
    # 4. START
    # ========================================================

    if exact_phrase_match(command, START_PHRASES):
        start_presentation()
        return True

    if safe_phrase_match(command, START_PHRASES, threshold=0.74):
        start_presentation()
        return True

    start_intent_words = {
        "start", "begin", "commence", "launch", "present",
        "presentation", "presenting", "started", "stark", "stat",
        "starr", "star"
    }

    if any(word in start_intent_words for word in command.split()):
        # A slide-number command has already been handled above.
        start_presentation()
        return True

    # ========================================================
    # 5. STOP
    # ========================================================

    if exact_phrase_match(command, STOP_PHRASES):
        stop_presentation()
        return True

    if safe_phrase_match(command, STOP_PHRASES, threshold=0.82):
        stop_presentation()
        return True

    stop_intent_words = {
        "stop", "end", "finish", "exit", "quit", "close",
        "stops", "stopped", "stopp", "shop", "stap"
    }

    if any(word in stop_intent_words for word in command.split()):
        stop_presentation()
        return True

    print(
        "HeritageAI: Command not recognized."
    )

    return None


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
        "HeritageAI: Loading voice model..."
    )

    if os.path.isdir(
        MODEL_PATH_INTERNAL
    ):

        model_path = (
            MODEL_PATH_INTERNAL
        )

    elif os.path.isdir(
        MODEL_PATH_LOCAL
    ):

        model_path = (
            MODEL_PATH_LOCAL
        )

    else:

        print()
        print(
            "HeritageAI: "
            "Voice model folder was not found."
        )

        print()
        print(
            "Expected:"
        )

        print(
            MODEL_PATH_INTERNAL
        )

        print(
            MODEL_PATH_LOCAL
        )

        return False

    try:

        model = vosk.Model(
            model_path
        )

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

        recognizer = (
            vosk.KaldiRecognizer(
                model,
                sample_rate
            )
        )

        recognizer.SetWords(
            True
        )

        # Ask Vosk for several plausible transcriptions.
        # The command engine can then recover when the first
        # transcription is slightly wrong.
        recognizer.SetMaxAlternatives(
            5
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
# MAIN
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
    print("Commands:")

    print(
        "  start presentation"
    )

    print(
        "  begin presentation"
    )

    print(
        "  number 3"
    )

    print(
        "  number three"
    )

    print(
        "  go to slide 3"
    )

    print(
        "  go to three"
    )

    print(
        "  next"
    )

    print(
        "  go forward"
    )

    print(
        "  previous"
    )

    print(
        "  go back"
    )

    print(
        "  stop"
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
        "HeritageAI: Testing microphone..."
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

                    data = audio_queue.get(
                        timeout=1
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

                        alternatives = result.get(
                            "alternatives",
                            []
                        )

                        if alternatives:

                            # Try the best alternative first.
                            # If it is empty, fall back to normal text.
                            candidate_texts = [
                                item.get("text", "")
                                for item in alternatives
                                if item.get("text", "")
                            ]
                            if candidate_texts:
                                for candidate in candidate_texts:
                                    if not candidate:
                                        continue
                                    command_result = process_command(candidate)
                                    if command_result is not None:
                                        running = command_result
                                        break
                            else:
                                text = result.get("text", "")
                                if text:
                                    running = process_command(text)

                        else:

                            text = result.get(
                                "text",
                                ""
                            )

                            if text:
                                running = process_command(text)

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
# HERITAGEAI START INTERFACE
# ============================================================

def launch_interface():

    import tkinter as tk
    from tkinter import filedialog, messagebox

    # --------------------------------------------------------
    # COLORS
    # --------------------------------------------------------

    BG = "#080808"
    PANEL = "#111111"
    PANEL_2 = "#161616"
    GOLD = "#D4AF37"
    GOLD_LIGHT = "#E8C95A"
    WHITE = "#F5F5F5"
    MUTED = "#929292"
    BORDER = "#292929"
    GREEN = "#5FCB81"

    # --------------------------------------------------------
    # WINDOW
    # --------------------------------------------------------

    root = tk.Tk()

    root.title("HeritageAI")
    root.geometry("820x560")
    root.minsize(760, 520)
    root.configure(bg=BG)
    root.resizable(True, True)

    # --------------------------------------------------------
    # CENTER WINDOW
    # --------------------------------------------------------

    root.update_idletasks()

    width = 820
    height = 560

    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    x = (screen_width - width) // 2
    y = (screen_height - height) // 2

    root.geometry(f"{width}x{height}+{x}+{y}")

    # --------------------------------------------------------
    # MAIN CONTAINER
    # --------------------------------------------------------

    main_frame = tk.Frame(
        root,
        bg=BG
    )

    main_frame.pack(
        fill="both",
        expand=True,
        padx=42,
        pady=32
    )

    # --------------------------------------------------------
    # HEADER
    # --------------------------------------------------------

    header = tk.Frame(
        main_frame,
        bg=BG
    )

    header.pack(
        fill="x"
    )

    brand = tk.Label(
        header,
        text="HERITAGE",
        bg=BG,
        fg=WHITE,
        font=("Arial", 22, "bold")
    )

    brand.pack(
        side="left"
    )

    ai_text = tk.Label(
        header,
        text="AI",
        bg=BG,
        fg=GOLD,
        font=("Arial", 22, "bold")
    )

    ai_text.pack(
        side="left"
    )

    version = tk.Label(
        header,
        text="  POWERPOINT VOICE CONTROL",
        bg=BG,
        fg=MUTED,
        font=("Arial", 9, "bold")
    )

    version.pack(
        side="left",
        padx=(10, 0),
        pady=(7, 0)
    )

    # --------------------------------------------------------
    # GOLD LINE
    # --------------------------------------------------------

    line = tk.Frame(
        main_frame,
        bg=GOLD,
        height=2
    )

    line.pack(
        fill="x",
        pady=(18, 30)
    )

    # --------------------------------------------------------
    # HERO SECTION
    # --------------------------------------------------------

    hero = tk.Frame(
        main_frame,
        bg=BG
    )

    hero.pack(
        fill="x"
    )

    hero_title = tk.Label(
        hero,
        text="Present With Your Voice.",
        bg=BG,
        fg=WHITE,
        font=("Arial", 30, "bold")
    )

    hero_title.pack(
        anchor="w"
    )

    hero_subtitle = tk.Label(
        hero,
        text=(
            "Control your PowerPoint presentation naturally "
            "while you present."
        ),
        bg=BG,
        fg=MUTED,
        font=("Arial", 11)
    )

    hero_subtitle.pack(
        anchor="w",
        pady=(8, 0)
    )

    # --------------------------------------------------------
    # PRESENTATION CARD
    # --------------------------------------------------------

    card = tk.Frame(
        main_frame,
        bg=PANEL,
        highlightbackground=BORDER,
        highlightthickness=1
    )

    card.pack(
        fill="x",
        pady=(30, 0)
    )

    card_inner = tk.Frame(
        card,
        bg=PANEL
    )

    card_inner.pack(
        fill="both",
        padx=24,
        pady=24
    )

    section_title = tk.Label(
        card_inner,
        text="PRESENTATION",
        bg=PANEL,
        fg=GOLD,
        font=("Arial", 10, "bold")
    )

    section_title.pack(
        anchor="w"
    )

    status = tk.Label(
        card_inner,
        text="No presentation selected",
        bg=PANEL,
        fg=WHITE,
        font=("Arial", 13, "bold"),
        anchor="w"
    )

    status.pack(
        fill="x",
        pady=(10, 3)
    )

    path_label = tk.Label(
        card_inner,
        text="Choose a PowerPoint file to get started.",
        bg=PANEL,
        fg=MUTED,
        font=("Arial", 9),
        anchor="w"
    )

    path_label.pack(
        fill="x"
    )

    # --------------------------------------------------------
    # BUTTON AREA
    # --------------------------------------------------------

    buttons = tk.Frame(
        card_inner,
        bg=PANEL
    )

    buttons.pack(
        fill="x",
        pady=(22, 0)
    )

    selected_file = {
        "path": None
    }

    # --------------------------------------------------------
    # BROWSE FUNCTION
    # --------------------------------------------------------

    def browse():

        file_path = filedialog.askopenfilename(
            title="Choose PowerPoint Presentation",
            filetypes=[
                (
                    "PowerPoint Presentation",
                    "*.pptx"
                ),
                (
                    "PowerPoint",
                    "*.ppt"
                ),
                (
                    "All Files",
                    "*.*"
                )
            ]
        )

        if not file_path:
            return

        selected_file["path"] = file_path

        filename = os.path.basename(
            file_path
        )

        status.config(
            text=filename,
            fg=WHITE
        )

        path_label.config(
            text=file_path,
            fg=MUTED
        )

        start_button.config(
            state="normal",
            bg=GOLD,
            fg="#080808"
        )

    # --------------------------------------------------------
    # START FUNCTION
    # --------------------------------------------------------

    def start_presentation():

        file_path = selected_file["path"]

        if not file_path:

            messagebox.showwarning(
                "HeritageAI",
                "Please choose a PowerPoint presentation first."
            )

            return

        start_button.config(
            state="disabled",
            text="Opening..."
        )

        status.config(
            text="Opening presentation...",
            fg=GOLD
        )

        root.update()

        try:

            global powerpoint
            global presentation

            file_path = os.path.abspath(
                os.path.normpath(
                    file_path
                )
            )

            if not os.path.isfile(
                file_path
            ):

                raise FileNotFoundError(
                    file_path
                )

            # ------------------------------------------------
            # SAME POWERPOINT CONNECTION AS YOUR WORKING UI
            # ------------------------------------------------

            powerpoint = (
                win32com.client.Dispatch(
                    "PowerPoint.Application"
                )
            )

            try:

                powerpoint.Visible = True

            except Exception:

                pass

            presentation = (
                powerpoint.Presentations.Open(
                    FileName=file_path,
                    ReadOnly=False,
                    Untitled=False,
                    WithWindow=True
                )
            )

            time.sleep(1)

            root.destroy()

            main()

        except Exception as error:

            start_button.config(
                state="normal",
                text="Start Presentation"
            )

            status.config(
                text="Could not open presentation",
                fg="#E06C75"
            )

            messagebox.showerror(
                "HeritageAI",
                "Could not open the presentation.\n\n"
                + str(error)
            )

    # --------------------------------------------------------
    # BROWSE BUTTON
    # --------------------------------------------------------

    browse_button = tk.Button(
        buttons,
        text="Choose Presentation",
        command=browse,
        bg=PANEL_2,
        fg=WHITE,
        activebackground="#202020",
        activeforeground=GOLD_LIGHT,
        relief="flat",
        bd=0,
        highlightthickness=1,
        highlightbackground=BORDER,
        font=("Arial", 10, "bold"),
        padx=18,
        pady=11,
        cursor="hand2"
    )

    browse_button.pack(
        side="left"
    )

    # --------------------------------------------------------
    # START BUTTON
    # --------------------------------------------------------

    start_button = tk.Button(
        buttons,
        text="Start Presentation",
        command=start_presentation,
        bg="#242424",
        fg="#777777",
        activebackground=GOLD_LIGHT,
        activeforeground="#080808",
        relief="flat",
        bd=0,
        font=("Arial", 10, "bold"),
        padx=20,
        pady=11,
        state="disabled",
        cursor="hand2"
    )

    start_button.pack(
        side="right"
    )

    # --------------------------------------------------------
    # VOICE COMMANDS
    # --------------------------------------------------------

    voice_section = tk.Frame(
        main_frame,
        bg=BG
    )

    voice_section.pack(
        fill="x",
        pady=(28, 0)
    )

    voice_title = tk.Label(
        voice_section,
        text="VOICE CONTROLS",
        bg=BG,
        fg=GOLD,
        font=("Arial", 9, "bold")
    )

    voice_title.pack(
        anchor="w"
    )

    command_frame = tk.Frame(
        voice_section,
        bg=BG
    )

    command_frame.pack(
        fill="x",
        pady=(12, 0)
    )

    commands = [
        ("NEXT", "Heritage, next"),
        ("PREVIOUS", "Heritage, previous"),
        ("SLIDE", "Heritage, slide 3"),
        ("START", "Heritage, start presentation"),
        ("STOP", "Heritage, stop")
    ]

    for title, command in commands:

        command_card = tk.Frame(
            command_frame,
            bg=PANEL,
            highlightbackground=BORDER,
            highlightthickness=1
        )

        command_card.pack(
            side="left",
            fill="both",
            expand=True,
            padx=(0, 8)
        )

        command_title = tk.Label(
            command_card,
            text=title,
            bg=PANEL,
            fg=WHITE,
            font=("Arial", 8, "bold")
        )

        command_title.pack(
            pady=(10, 2)
        )

        command_text = tk.Label(
            command_card,
            text=command,
            bg=PANEL,
            fg=MUTED,
            font=("Arial", 8)
        )

        command_text.pack(
            pady=(0, 10)
        )

    # --------------------------------------------------------
    # FOOTER
    # --------------------------------------------------------

    footer = tk.Frame(
        main_frame,
        bg=BG
    )

    footer.pack(
        fill="x",
        side="bottom",
        pady=(20, 0)
    )

    footer_left = tk.Label(
        footer,
        text="●  Voice control ready",
        bg=BG,
        fg=GREEN,
        font=("Arial", 9, "bold")
    )

    footer_left.pack(
        side="left"
    )

    footer_right = tk.Label(
        footer,
        text="HeritageAI  •  Present With Your Voice",
        bg=BG,
        fg=MUTED,
        font=("Arial", 8)
    )

    footer_right.pack(
        side="right"
    )

    # --------------------------------------------------------
    # START APPLICATION
    # --------------------------------------------------------

    root.mainloop()


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":

    launch_interface()