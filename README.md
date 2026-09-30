# DSolve-2025 — Universal Game Accessibility Tool

Hands-free gaming input that works with any game, no game modifications needed.
It listens to voice commands and tracks head tilt via webcam, converting them
into keyboard input at the OS level (Windows).

## Team

- Janbaas Jamal K K
- Hani Muhamed

## Project Idea

Many games lack built-in accessibility features. This tool provides:

- Voice-controlled inputs for hands-free gaming (offline speech recognition).
- Head-tilt controls via webcam for easy interaction.

Planned but not yet implemented: color filters for color blindness.

## Key Features

- 🎙️ Voice commands (`spp.py:448-453`): up, down, left, right, jump (space),
  move/start (W), attack (X), punch (E) — plus switch, enter, tab, escape in
  the key map. Fuzzy-matched with difflib, so near-misses still register.
- 🧠 Head-tilt steering (`spp.py:56-248`): MediaPipe FaceMesh measures roll
  angle between the ears; tilting past the threshold (default 30°) presses
  `A` (left) or `D` (right). Includes smoothing and a live camera preview
  (press `Q` in the preview window to stop it).
- 🪟 Game-friendly key injection (`spp.py:22-54`): sends low-level key events
  via Win32 (`keybd_event`) in addition to `keyboard`, so keys register in
  games that ignore synthetic input.

## Tech Stack

| Technology | Version | Purpose |
|------------|---------|---------|
| Python | 3.12.10 (64-bit Windows) | Runtime (see `.python-version`) |
| Vosk | 0.3.45 | Offline speech recognition (`Vosk/` model dir) |
| sounddevice | 0.5.6 | Microphone capture |
| MediaPipe | 0.10.14 | FaceMesh head-tilt tracking |
| opencv-python | 4.8.1.78 | Webcam capture + preview |
| numpy | 1.26.4 | Angle math (must stay 1.x — see note) |
| keyboard | 0.13.5 | Key press simulation |
| pywin32 | 312 | Low-level Windows key events |

> Version notes: MediaPipe must stay at 0.10.14 — newer releases removed the
> `mediapipe.solutions.face_mesh` API this project uses. NumPy must stay
> 1.x — the pinned OpenCV/MediaPipe builds crash on NumPy 2.x. Python 3.14
> has no MediaPipe wheels, so use 3.12.

## Setup Instructions

### Prerequisites

- Windows 10/11 (64-bit), Python 3.12
- Microphone + webcam
- Vosk English model extracted to `Vosk/` (the repo includes
  `vosk-model-small-en-us-0.15.zip` — extract it so `Vosk/am`, `Vosk/conf`,
  `Vosk/graph`, `Vosk/ivector` exist at that path)
- Run the terminal as Administrator for key presses to reach games

### Installation

```powershell
py -3.12 -m venv venv
.\venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

### Running the Project

```powershell
.\venv\Scripts\Activate.ps1
python spp.py
```

Say a command (e.g. "jump", "left") or tilt your head left/right. A camera
preview window opens — press `Q` in it to stop tracking, or `Ctrl+C` in the
terminal to stop everything.

## How to Contribute

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/your-feature`)
3. Commit your changes (`git commit -m 'Add some feature'`)
4. Push to the branch (`git push origin feature/your-feature`)
5. Open a Pull Request

Ideas: color-blindness filters, configurable voice→key bindings, per-game
profiles, Linux support (replace `pywin32` input layer).

## Acknowledgments

- [Vosk offline speech recognition](https://alphacephei.com/vosk/)
- [MediaPipe FaceMesh](https://google.github.io/mediapipe/solutions/face_mesh)
- [OpenCV](https://opencv.org/)
