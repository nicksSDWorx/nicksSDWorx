# Cosmic Catcher

A small arcade minigame written in pure-stdlib Python (tkinter, no extra
deps).  Runs natively on Windows.

![cosmic catcher](https://img.shields.io/badge/platform-windows-blue)

## Two ways to play

### 1. Quickest - double-click `run.bat`

Requires Python 3 installed on the PC (any version from python.org with
the default options is fine; tkinter ships with it).

```
run.bat
```

That's it.

### 2. Make a real standalone `.exe`

Run this once:

```
build_exe.bat
```

It installs PyInstaller and builds `dist\CosmicCatcher.exe` - a single
self-contained executable.  You can then move that `.exe` to any
Windows machine and double-click it; Python does **not** need to be
installed on the target machine.

## How to play

You pilot a tiny ship at the bottom of the screen.  Things rain from
space:

| Icon | Item     | Effect                          |
|------|----------|---------------------------------|
| ★    | Star     | +10 points                      |
| ◆    | Gem      | +50 points (rare)               |
| ●    | Bomb     | -1 life - **dodge!**            |
| S    | Shield   | 2 seconds of invulnerability    |
| %    | Slow-mo  | 3 seconds of bullet-time        |

Catch goodies in a row to build a **combo multiplier** (x2 → x5).
Missing a star or gem, or eating a bomb, breaks the combo.  Every 500
points the **wave** increases - everything falls faster and bombs get
more frequent.

You start with 3 lives.

### Controls

| Key                     | Action                |
|-------------------------|-----------------------|
| ← / →   or   A / D      | Move                  |
| Space (hold)            | Brake (half speed)    |
| P                       | Pause / resume        |
| R                       | Restart after game over |
| Esc                     | Quit                  |

Have fun.
