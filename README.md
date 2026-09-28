# HeritageAI

### Present With Your Voice.

**HeritageAI** is a voice-controlled PowerPoint presentation assistant for Windows that lets presenters navigate their slides hands-free using spoken commands.

Instead of reaching for a keyboard or mouse during a presentation, you can use your voice to start the presentation, move between slides, jump directly to a slide, find slides by topic, and stop the presentation.

---

## 🎥 See HeritageAI in Action

> Demo video coming soon.

**Website:**
https://herrie23-dev.github.io/HeritageAI/

**Latest Release:**
Check the [Releases](../../releases) page for the latest Windows version.

---

## ✨ What HeritageAI Can Do

### 🎙️ Voice-Controlled Navigation

Control Microsoft PowerPoint using spoken commands such as:

* `Next`
* `Go forward`
* `Show the next slide`
* `Previous`
* `Go back`
* `Move backward`

### 🔢 Jump Directly to a Slide

Instead of moving through slides one at a time:

* `Go to slide 3`
* `Show slide three`
* `Take me to slide five`

### 🔎 Find Slides by Topic

HeritageAI can examine the text contained in your presentation and help you navigate to a slide based on what you want to discuss.

For example:

* `Show the methodology`
* `Take me to the conclusion`
* `Where did I discuss networking?`

This allows presenters to navigate presentations based on their content, rather than only relying on slide numbers.

### 🧠 Flexible Speech Interpretation

HeritageAI processes spoken commands through speech recognition, text normalization, correction rules, fuzzy matching, and presentation-aware command handling.

This helps it interpret natural variations in how a presenter may phrase a command.

### 🔌 Works Offline

Speech recognition is powered by **Vosk**, allowing the core voice-recognition system to operate locally without requiring an internet connection.

### 🖥️ Windows Desktop Application

HeritageAI integrates directly with Microsoft PowerPoint on Windows and provides a dedicated desktop interface for selecting and controlling presentations.

---

## 🎤 Example Voice Commands

| Purpose           | Example                          |
| ----------------- | -------------------------------- |
| Start             | `Start presentation`             |
| Next slide        | `Next`                           |
| Next slide        | `Show the next slide`            |
| Previous slide    | `Previous`                       |
| Previous slide    | `Go back`                        |
| Direct navigation | `Go to slide 3`                  |
| Direct navigation | `Show slide three`               |
| Topic navigation  | `Show the methodology`           |
| Topic navigation  | `Take me to the conclusion`      |
| Topic navigation  | `Where did I discuss networking` |
| Stop              | `Stop presentation`              |

---

## ⚙️ How It Works

HeritageAI follows a voice-to-command pipeline:

```text
🎤 Microphone
      ↓
🗣️ Vosk Speech Recognition
      ↓
🧹 Speech Normalization & Correction
      ↓
🧠 Command Interpretation
      ↓
📑 Presentation / Topic Analysis
      ↓
🖥️ Microsoft PowerPoint
```

The system combines speech recognition with command interpretation and presentation text analysis to turn spoken instructions into PowerPoint actions.

---

## 🚀 Getting Started

### Option 1 — Download the Windows Release

If you simply want to use HeritageAI, download the latest packaged version from:

**[GitHub Releases](../../releases)**

The packaged application is intended for Windows users who already have Microsoft PowerPoint installed.

### Requirements

* Windows
* Microsoft PowerPoint
* A working microphone
* A PowerPoint presentation (`.pptx` or `.ppt`)

---

## 🛠️ Run From Source

### 1. Clone the repository

```bash
git clone https://github.com/Herrie23-dev/HeritageAI.git
cd HeritageAI
```

### 2. Install the required Python packages

```bash
pip install vosk sounddevice pywin32
```

### 3. Download the Vosk model

HeritageAI currently uses:

```text
vosk-model-small-en-us-0.15
```

Place the model where the application can locate it, including the expected `_internal` location when using the packaged application.

### 4. Run HeritageAI

```bash
python main.py
```

Make sure Microsoft PowerPoint is installed on the computer before using presentation control.

---

## 🧩 Project Structure

```text
HeritageAI/
│
├── .github/
│   └── workflows/
│
├── vosk-model-small-en-us-0.15/
│
├── index.html
├── privacy.html
├── sitemap.xml
├── main.py
├── README.md
└── ...
```

---

## 🧠 Technology

HeritageAI currently uses:

* **Python** — core application logic
* **Vosk** — offline speech recognition
* **SoundDevice** — microphone/audio input
* **Tkinter** — desktop user interface
* **PyWin32 / COM** — Microsoft PowerPoint integration
* **PowerPoint** — presentation control

---

## 🗺️ Roadmap

HeritageAI is an evolving project.

Areas being explored include:

* More intelligent presentation assistance
* Improved speech interpretation
* More flexible voice commands
* Better presentation context understanding
* Expanded presentation navigation
* Improved user interface
* Additional accessibility features
* Broader language and speech support
* More intelligent AI-assisted presentation features

Features on the roadmap are not necessarily available in the current release.

---

## 🤝 Contributing

HeritageAI is being developed as an evolving project, and contributions, suggestions, bug reports, and feature ideas are welcome.

If you find a problem or have an idea:

1. Open an issue.
2. Describe the problem or proposed feature.
3. Include steps to reproduce the issue where applicable.
4. Provide relevant system information when reporting bugs.

More detailed contribution guidelines will be added as the project grows.

---

## 🔐 Privacy

HeritageAI is designed around local speech recognition using Vosk.

For more information about the project's privacy approach, see:

**[Privacy Policy](privacy.html)**

---

## 👨🏽‍💻 About the Developer

HeritageAI was created and developed by **Bolarinwa Heritage Oluwanimilo**, a Computer Engineering student interested in artificial intelligence, machine learning, and practical technology solutions.

The project began from a simple idea:

> **What if you could control your presentation without touching your laptop?**

HeritageAI is an ongoing attempt to turn that idea into a practical presentation tool.

---

## 🌍 Vision

### Present with your voice. Focus on your audience.

HeritageAI aims to make presentations more natural, accessible, and hands-free by allowing presenters to interact with their presentation using their voice.

---

## 📌 Project Links

* **Website:** https://herrie23-dev.github.io/HeritageAI/
* **Repository:** https://github.com/Herrie23-dev/HeritageAI
* **Releases:** [View Releases](../../releases)
* **Issues:** [Report an Issue](../../issues)

---

## ⭐ Support the Project

If you find HeritageAI interesting or useful:

* ⭐ Star the repository
* 🐛 Report bugs
* 💡 Suggest features
* 🎥 Share the project
* 🤝 Contribute to the project

Every genuine interaction helps the project reach more people.

---

**HeritageAI — Present With Your Voice.**
