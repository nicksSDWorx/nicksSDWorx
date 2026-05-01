"""
PAYDAY PANIC
============
It's Friday, 4 minutes to 5pm.  A stack of timesheets just landed on
your desk and the conveyor belt won't stop.  Sort them all before
payroll closes.

Each card slides in from the right.  Read the data, then sort the
*front-most* (leftmost) card into one of three trays:

    A  /  1   APPROVE      hours sane (1-60) and status OK
    H  /  2   HOLD         status PENDING - awaiting manager sign-off
    R  /  3   REJECT       hours 0, hours > 60, or status ERROR

Specials
--------
    BONUS CHECK   always APPROVE  -> +100
    AUDITOR       always REJECT   -> approving them is a disaster
    COFFEE  ☕    any action eats it and slows the belt for 4 sec

Wrong action  -> -3 sec on the clock and combo broken.
Card escapes  -> -2 sec and combo broken.
3 correct in a row starts a combo multiplier (up to x4).

You have 90 seconds.  Final number is your payroll score.

Controls
--------
    A / 1   Approve              R / 3   Reject
    H / 2   Hold                 P       Pause
    Space   Cycle to next card   Esc     Quit
                                 Enter   Restart after game over
"""

import random
import time
import tkinter as tk

W, H = 760, 540
FPS = 60
FRAME_MS = 1000 // FPS

CARD_W, CARD_H = 150, 110
CONVEYOR_Y = 170
CONVEYOR_BOTTOM = CONVEYOR_Y + CARD_H + 20
DEFAULT_SPEED = 1.05      # px per "60fps frame"
SPAWN_INTERVAL = 70       # frames between cards (gets faster)
GAME_LENGTH = 90.0        # seconds

FIRST_NAMES = ["Alex", "Sam", "Robin", "Jordan", "Taylor", "Casey", "Morgan",
               "Jamie", "Riley", "Quinn", "Avery", "Sky", "Drew", "Pat",
               "Kim", "Lee", "Max", "Noor", "Ola", "Reese"]
LAST_NAMES  = ["Chen", "Patel", "Garcia", "Smith", "Nguyen", "Kowalski",
               "Okafor", "Müller", "Rossi", "Park", "Singh", "Dubois",
               "van Dijk", "Andersen", "Costa", "Tanaka"]

# (kind, weight)
NORMAL_WEIGHTS = [
    ("ok",       55),    # hours 1..60 + OK
    ("overtime", 12),    # hours 61..90 -> reject
    ("zero",      8),    # hours 0      -> reject
    ("error",    10),    # status ERROR -> reject
    ("pending",  15),    # status PENDING -> hold
]
SPECIAL_WEIGHTS = [
    ("bonus",   6),
    ("auditor", 5),
    ("coffee",  4),
]

CORRECT = {
    "ok":       "approve",
    "overtime": "reject",
    "zero":     "reject",
    "error":    "reject",
    "pending":  "hold",
    "bonus":    "approve",
    "auditor":  "reject",
}

TRAY_COLORS = {
    "approve": "#4cd97a",
    "hold":    "#ffc94a",
    "reject":  "#ff5a5a",
}
TRAY_LABEL = {
    "approve": "APPROVE  [A/1]",
    "hold":    "HOLD     [H/2]",
    "reject":  "REJECT   [R/3]",
}


def random_name():
    return f"{random.choice(FIRST_NAMES)} {random.choice(LAST_NAMES)}"


def pick_kind():
    if random.random() < 0.18:
        pool = SPECIAL_WEIGHTS
    else:
        pool = NORMAL_WEIGHTS
    total = sum(w for _, w in pool)
    r = random.uniform(0, total)
    acc = 0
    for k, w in pool:
        acc += w
        if r <= acc:
            return k
    return pool[-1][0]


def make_card_data(kind):
    name = random_name()
    rate = random.choice([14, 18, 22, 28, 35, 42, 55])
    if kind == "ok":
        hours, status = random.randint(20, 60), "OK"
    elif kind == "overtime":
        hours, status = random.randint(61, 99), "OK"
    elif kind == "zero":
        hours, status = 0, "OK"
    elif kind == "error":
        hours, status = random.randint(20, 60), "ERROR"
    elif kind == "pending":
        hours, status = random.randint(20, 60), "PENDING"
    elif kind == "bonus":
        name = "BONUS CHECK"
        hours, status = random.randint(40, 60), "BONUS"
    elif kind == "auditor":
        name = "AUDITOR  (do NOT pay)"
        hours, status = random.randint(80, 200), "AUDIT"
    elif kind == "coffee":
        name = "COFFEE  (perk)"
        hours, status = 0, "COFFEE"
    else:
        hours, status = 40, "OK"
    return {"name": name, "hours": hours, "rate": rate, "status": status, "kind": kind}


class Card:
    def __init__(self, canvas, x, data):
        self.data = data
        self.x = x
        self.y = CONVEYOR_Y
        self.flying = False           # set when sorted - card flies to tray
        self.flight_vx = 0.0
        self.flight_vy = 0.0
        self.flight_life = 0.0
        self.dead = False
        self.ids = []
        self._draw(canvas)

    def _draw(self, canvas):
        d = self.data
        kind = d["kind"]
        if kind == "bonus":
            bg, border = "#1f3a1a", "#7bff8a"
        elif kind == "auditor":
            bg, border = "#3a1010", "#ff7777"
        elif kind == "coffee":
            bg, border = "#3a2a14", "#d8a86c"
        else:
            bg, border = "#181d2e", "#3d4775"
        x, y = self.x, self.y
        self.ids.append(canvas.create_rectangle(
            x, y, x + CARD_W, y + CARD_H,
            fill=bg, outline=border, width=2,
        ))
        # status strip
        strip_color = {
            "OK": "#4cd97a", "ERROR": "#ff5a5a", "PENDING": "#ffc94a",
            "BONUS": "#7bff8a", "AUDIT": "#ff5a5a", "COFFEE": "#d8a86c",
        }[d["status"]]
        self.ids.append(canvas.create_rectangle(
            x, y, x + CARD_W, y + 14, fill=strip_color, outline=""))
        self.ids.append(canvas.create_text(
            x + CARD_W / 2, y + 7, text=d["status"],
            fill="#0c0c0c", font=("Consolas", 8, "bold")))
        # name
        name_text = d["name"]
        if len(name_text) > 18:
            name_text = name_text[:17] + "…"
        self.ids.append(canvas.create_text(
            x + CARD_W / 2, y + 32, text=name_text,
            fill="#ffffff", font=("Consolas", 10, "bold")))
        # hours / rate
        if kind == "coffee":
            self.ids.append(canvas.create_text(
                x + CARD_W / 2, y + 64, text="☕  break time?",
                fill="#d8a86c", font=("Consolas", 11)))
        elif kind == "auditor":
            self.ids.append(canvas.create_text(
                x + CARD_W / 2, y + 56, text="claims " + str(d["hours"]) + " h",
                fill="#ff9a9a", font=("Consolas", 10)))
            self.ids.append(canvas.create_text(
                x + CARD_W / 2, y + 76, text="@ $" + str(d["rate"]) + "/h",
                fill="#ff9a9a", font=("Consolas", 10)))
        else:
            self.ids.append(canvas.create_text(
                x + CARD_W / 2, y + 56, text=f"hours: {d['hours']}",
                fill="#cfd6f5", font=("Consolas", 11)))
            self.ids.append(canvas.create_text(
                x + CARD_W / 2, y + 76, text=f"rate:  ${d['rate']}/h",
                fill="#cfd6f5", font=("Consolas", 11)))
        # hint icon
        self.ids.append(canvas.create_text(
            x + CARD_W / 2, y + CARD_H - 14,
            text={"approve": "✓", "hold": "⏸", "reject": "✕"}[CORRECT[kind]],
            fill="#2a2f44", font=("Consolas", 9)))

    def move(self, canvas, dx, dy):
        self.x += dx
        self.y += dy
        for i in self.ids:
            canvas.move(i, dx, dy)

    def destroy(self, canvas):
        for i in self.ids:
            canvas.delete(i)
        self.ids = []
        self.dead = True


class Game:
    def __init__(self, root):
        self.root = root
        root.title("Payday Panic")
        root.resizable(False, False)
        root.configure(bg="#0b0d18")
        self.canvas = tk.Canvas(root, width=W, height=H, bg="#0b0d18", highlightthickness=0)
        self.canvas.pack()

        # Office wallpaper - faint grid lines like a spreadsheet
        for gx in range(0, W, 40):
            self.canvas.create_line(gx, 60, gx, H - 110, fill="#11142a")
        for gy in range(60, H - 110, 40):
            self.canvas.create_line(0, gy, W, gy, fill="#11142a")

        # Conveyor belt area
        self.canvas.create_rectangle(0, CONVEYOR_Y - 16, W, CONVEYOR_BOTTOM,
                                     fill="#1a1d2e", outline="")
        for tx in range(0, W, 30):
            self.canvas.create_line(tx, CONVEYOR_Y - 16, tx + 12, CONVEYOR_Y - 16,
                                    fill="#3d4775", width=2)
            self.canvas.create_line(tx, CONVEYOR_BOTTOM, tx + 12, CONVEYOR_BOTTOM,
                                    fill="#3d4775", width=2)
        # Left wall (cards passing through = miss)
        self.canvas.create_line(0, CONVEYOR_Y - 30, 0, CONVEYOR_BOTTOM + 10,
                                fill="#ff5a5a", width=2)

        # Trays at the bottom
        self.tray_centers = {}
        tray_y = H - 70
        for i, action in enumerate(("approve", "hold", "reject")):
            cx = (i + 1) * (W / 4)
            self.tray_centers[action] = (cx, tray_y)
            color = TRAY_COLORS[action]
            self.canvas.create_rectangle(
                cx - 110, tray_y - 36, cx + 110, tray_y + 36,
                fill="#101326", outline=color, width=3)
            self.canvas.create_text(cx, tray_y - 14, text=TRAY_LABEL[action],
                                    fill=color, font=("Consolas", 11, "bold"))
            self.canvas.create_text(cx, tray_y + 12,
                                    text={"approve": "hours OK & status OK",
                                          "hold":    "PENDING",
                                          "reject":  "0 / overtime / ERROR"}[action],
                                    fill="#7c84a8", font=("Consolas", 8))

        # HUD
        self.hud_clock = self.canvas.create_text(W / 2, 22, anchor="c", fill="#ffd54a",
                                                 font=("Consolas", 18, "bold"), text="")
        self.hud_score = self.canvas.create_text(14, 14, anchor="nw", fill="#7af6ff",
                                                 font=("Consolas", 13, "bold"), text="")
        self.hud_combo = self.canvas.create_text(W - 14, 14, anchor="ne", fill="#7bff8a",
                                                 font=("Consolas", 13, "bold"), text="")
        self.hud_status = self.canvas.create_text(W / 2, H / 2, anchor="c", fill="#ffffff",
                                                  font=("Consolas", 20, "bold"), text="",
                                                  justify="center")
        self.hud_help = self.canvas.create_text(W / 2, H - 8, anchor="s", fill="#4d567a",
                                                font=("Consolas", 9),
                                                text="A approve   H hold   R reject   Space cycle   P pause   Esc quit")

        # Floating toast (e.g. "+15", "WRONG -3s")
        self.toasts = []          # list of (id, vy, ttl)

        root.bind("<KeyPress>", self._on_key)
        root.protocol("WM_DELETE_WINDOW", root.destroy)

        self.reset()
        self.last_t = time.perf_counter()
        self.root.after(FRAME_MS, self.tick)

    # --- input ---
    def _on_key(self, e):
        k = e.keysym.lower()
        if k == "escape":
            self.root.destroy(); return
        if self.game_over:
            if k == "return":
                self.reset()
            return
        if k == "p":
            self.paused = not self.paused
            self.canvas.itemconfigure(self.hud_status, text="PAUSED" if self.paused else "")
            return
        if self.paused:
            return
        if k in ("a", "1"):     self._sort("approve")
        elif k in ("h", "2"):   self._sort("hold")
        elif k in ("r", "3"):   self._sort("reject")
        elif k == "space":      self._cycle()

    # --- state ---
    def reset(self):
        for c in getattr(self, "cards", []):
            c.destroy(self.canvas)
        for t in getattr(self, "toasts", []):
            self.canvas.delete(t[0])
        self.cards = []
        self.toasts = []
        self.score = 0
        self.combo = 0
        self.best_combo = 0
        self.processed = 0
        self.correct = 0
        self.spawn_timer = 0.0
        self.spawn_interval = SPAWN_INTERVAL
        self.speed = DEFAULT_SPEED
        self.slow_until = 0.0
        self.time_left = GAME_LENGTH
        self.paused = False
        self.game_over = False
        self.active_idx = 0       # which card is currently selected
        self.canvas.itemconfigure(self.hud_status, text="")
        self._refresh_hud()

    def _refresh_hud(self):
        m = int(self.time_left) // 60
        s = int(self.time_left) % 60
        self.canvas.itemconfigure(self.hud_clock, text=f"{m:01d}:{s:02d}")
        self.canvas.itemconfigure(self.hud_score, text=f"PAYROLL  ${self.score}")
        if self.combo >= 3:
            self.canvas.itemconfigure(self.hud_combo, text=f"COMBO x{self._mult()}  ({self.combo})")
        else:
            self.canvas.itemconfigure(self.hud_combo, text="")

    def _mult(self):
        if self.combo >= 12: return 4
        if self.combo >= 7:  return 3
        if self.combo >= 3:  return 2
        return 1

    # --- main loop ---
    def tick(self):
        now = time.perf_counter()
        dt_real = now - self.last_t
        self.last_t = now
        dt = max(0.0, min(3.0, dt_real * 60.0))

        if not self.paused and not self.game_over:
            self._update(dt, now)
        self.root.after(FRAME_MS, self.tick)

    def _update(self, dt, now):
        # clock
        self.time_left -= dt / 60.0
        if self.time_left <= 0:
            self.time_left = 0
            self._end_game()
            return

        # difficulty: belt slowly speeds up, spawn rate increases.
        progress = 1.0 - (self.time_left / GAME_LENGTH)
        speed = DEFAULT_SPEED + progress * 1.4
        spawn_int = max(28.0, SPAWN_INTERVAL - progress * 36.0)
        if now < self.slow_until:
            speed *= 0.45

        # spawn
        self.spawn_timer += dt
        # Avoid stacking - require last card to have moved away enough.
        can_spawn = True
        if self.cards:
            last = self.cards[-1]
            if not last.flying and last.x > W - CARD_W - 12:
                can_spawn = False
        if can_spawn and self.spawn_timer >= spawn_int:
            self.spawn_timer = 0.0
            self._spawn()

        # move cards
        keep = []
        for c in self.cards:
            if c.flying:
                c.move(self.canvas, c.flight_vx * dt, c.flight_vy * dt)
                c.flight_life -= dt / 60.0
                if c.flight_life <= 0:
                    c.destroy(self.canvas); continue
                keep.append(c); continue
            c.move(self.canvas, -speed * dt, 0)
            if c.x + CARD_W < 0:
                # Missed!
                self._missed(c)
                c.destroy(self.canvas)
                continue
            keep.append(c)
        self.cards = keep

        # clamp active selection
        live = [c for c in self.cards if not c.flying]
        if not live:
            self.active_idx = 0
        else:
            self.active_idx = max(0, min(self.active_idx, len(live) - 1))
            self._draw_selector(live[self.active_idx])

        # toasts
        nt = []
        for tid, vy, ttl in self.toasts:
            self.canvas.move(tid, 0, vy * dt)
            ttl -= dt / 60.0
            if ttl > 0:
                nt.append((tid, vy, ttl))
            else:
                self.canvas.delete(tid)
        self.toasts = nt

        self._refresh_hud()

    def _draw_selector(self, card):
        if not hasattr(self, "_selector_id") or self._selector_id is None:
            self._selector_id = self.canvas.create_rectangle(
                card.x - 4, card.y - 4, card.x + CARD_W + 4, card.y + CARD_H + 4,
                outline="#7af6ff", width=3)
        else:
            self.canvas.coords(self._selector_id,
                               card.x - 4, card.y - 4,
                               card.x + CARD_W + 4, card.y + CARD_H + 4)
            self.canvas.itemconfigure(self._selector_id, outline="#7af6ff")
        self.canvas.tag_raise(self._selector_id)

    def _spawn(self):
        kind = pick_kind()
        data = make_card_data(kind)
        x = W + 10
        c = Card(self.canvas, x, data)
        self.cards.append(c)

    def _live_cards(self):
        return [c for c in self.cards if not c.flying]

    def _cycle(self):
        live = self._live_cards()
        if not live: return
        self.active_idx = (self.active_idx + 1) % len(live)

    def _sort(self, action):
        live = self._live_cards()
        if not live:
            return
        # Use the *front-most* card (smallest x).
        card = min(live, key=lambda c: c.x)
        kind = card.data["kind"]

        if kind == "coffee":
            # Any action drinks the coffee.
            self.slow_until = time.perf_counter() + 4.0
            self._toast(card.x + CARD_W / 2, card.y, "☕  SLOW MO  4s", "#d8a86c")
            self._send_to_tray(card, action, "#d8a86c")
            return

        correct = CORRECT[kind]

        if action == correct:
            mult = self._mult()
            base = {"approve": 25, "hold": 20, "reject": 20}[action]
            if kind == "bonus":
                base = 100
            gained = base * mult
            self.score += gained
            self.combo += 1
            self.best_combo = max(self.best_combo, self.combo)
            self.correct += 1
            self.processed += 1
            self._toast(card.x + CARD_W / 2, card.y,
                        f"+${gained}" + (f"  x{mult}" if mult > 1 else ""),
                        TRAY_COLORS[action])
            self._send_to_tray(card, action, TRAY_COLORS[action])
        else:
            penalty = 5.0 if kind == "auditor" else 3.0
            self.time_left = max(0, self.time_left - penalty)
            self.combo = 0
            self.processed += 1
            self._toast(card.x + CARD_W / 2, card.y,
                        f"WRONG  -{int(penalty)}s", "#ff5a5a")
            self._send_to_tray(card, action, "#ff5a5a")

    def _send_to_tray(self, card, action, flash):
        cx, cy = self.tray_centers[action]
        # ballistic flight to the tray
        dx = (cx - (card.x + CARD_W / 2)) / 30.0
        dy = (cy - (card.y + CARD_H / 2)) / 30.0
        card.flight_vx = dx
        card.flight_vy = dy
        card.flight_life = 0.5
        card.flying = True
        # selector flashes the chosen tray
        self.canvas.itemconfigure(self._selector_id, outline=flash)

    def _missed(self, card):
        # Coffee escaping is fine, no penalty.
        if card.data["kind"] == "coffee":
            return
        self.time_left = max(0, self.time_left - 2.0)
        self.combo = 0
        self.processed += 1
        self._toast(20, CONVEYOR_Y + CARD_H / 2, "MISS  -2s", "#ff5a5a")

    def _toast(self, x, y, msg, color):
        tid = self.canvas.create_text(x, y, text=msg, fill=color,
                                      font=("Consolas", 13, "bold"))
        self.toasts.append((tid, -0.5, 1.0))

    def _end_game(self):
        self.game_over = True
        acc = (self.correct / self.processed * 100) if self.processed else 0
        self.canvas.itemconfigure(
            self.hud_status,
            text=("📊  PAYROLL CLOSED  📊\n"
                  f"final:  ${self.score}\n"
                  f"processed:  {self.processed}\n"
                  f"accuracy:  {acc:.0f}%\n"
                  f"best combo:  x{self._mult_for(self.best_combo)} ({self.best_combo})\n"
                  "press Enter to retry"))

    @staticmethod
    def _mult_for(combo):
        if combo >= 12: return 4
        if combo >= 7:  return 3
        if combo >= 3:  return 2
        return 1


def main():
    root = tk.Tk()
    Game(root)
    root.mainloop()


if __name__ == "__main__":
    main()
