# Payday Panic

A payroll-themed arcade minigame.  It's Friday afternoon and a stack
of timesheets just hit your conveyor belt.  Sort each one into the
right tray before payroll closes at 5pm.

Pure-stdlib Python (tkinter, no extra deps).  Runs natively on
Windows.

## Run it

Quickest:

```
run_payday.bat
```

Build a real standalone `.exe`:

```
build_payday_exe.bat
```

That writes `dist\PaydayPanic.exe`, which runs on any Windows machine
without Python installed.

## How to play

Cards slide in from the right.  Each one shows an employee, their
hours, their hourly rate, and a status flag.  Sort the **front-most**
card (the one closest to falling off the left edge — the cyan
selector frame highlights it):

| Tray         | When to use                              | Reward |
|--------------|------------------------------------------|--------|
| **APPROVE**  | hours sane (1–60) **and** status `OK`    | +$25   |
| **HOLD**     | status `PENDING`                         | +$20   |
| **REJECT**   | hours `0`, hours `> 60`, or status `ERROR` | +$20 |

A correct call in a row builds a **combo multiplier** up to **x4**.

### Specials

| Card           | Correct action          | Why                                  |
|----------------|-------------------------|--------------------------------------|
| **BONUS CHECK** (green) | APPROVE        | +$100 instant payout                 |
| **AUDITOR** (red)       | REJECT         | Approving them is a *huge* mistake (-5s) |
| **COFFEE** ☕            | any action     | Drink it → conveyor slows for 4 sec  |

### Penalties

- Wrong action → **−3 sec** on the clock (−5 for auditor) and combo broken
- Card escapes off the left edge → **−2 sec** and combo broken

You have **90 seconds**.  Final score is your processed payroll total.

### Controls

| Key             | Action                                |
|-----------------|---------------------------------------|
| `A` / `1`       | Approve                               |
| `H` / `2`       | Hold                                  |
| `R` / `3`       | Reject                                |
| `Space`         | Cycle the cyan selector to next card  |
| `P`             | Pause / resume                        |
| `Enter`         | Restart after game over               |
| `Esc`           | Quit                                  |

Good luck and watch out for the auditor.
