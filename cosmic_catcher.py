"""
COSMIC CATCHER
==============
A creative little arcade game.

You pilot a tiny ship at the bottom of the screen.  Stuff rains from
space:

    *   stars       +10 points       (yellow)
    @   gems        +50 points       (cyan, rare)
    #   bombs       -1 life          (red, avoid!)
    S   shields     2 sec invuln     (green, rare)
    %   slow-mo     3 sec time slow  (purple, rare)

Catch enough goodies in a row and the combo multiplier kicks in.
Every 500 points the wave levels up - things fall faster and more
often.  You start with 3 lives.

Controls
--------
    Left / Right arrows  or  A / D   - move
    Space                            - hold to brake (half speed)
    P                                - pause / resume
    R                                - restart after game over
    Esc                              - quit

Pure-stdlib Python (tkinter only).  Runs on any Windows machine that
has Python installed - just double-click run.bat.  See build_exe.bat
to bake it into a real one-file .exe with PyInstaller.
"""

import random
import time
import tkinter as tk

W, H = 720, 540
FPS = 60
FRAME_MS = 1000 // FPS

PLAYER_W, PLAYER_H = 54, 22
PLAYER_SPEED = 7.0

ITEM_TYPES = {
    "star":   {"color": "#ffd54a", "radius": 9,  "score":  10, "weight": 60, "shape": "star"},
    "gem":    {"color": "#4ad9ff", "radius": 10, "score":  50, "weight":  8, "shape": "diamond"},
    "bomb":   {"color": "#ff4a4a", "radius": 11, "score":   0, "weight": 22, "shape": "bomb"},
    "shield": {"color": "#6cff6c", "radius": 10, "score":   0, "weight":  5, "shape": "ring"},
    "slow":   {"color": "#c46cff", "radius": 10, "score":   0, "weight":  5, "shape": "ring"},
}


def weighted_pick(level):
    # Bombs get more common with level; goodies stay roughly the same.
    weights = []
    keys = list(ITEM_TYPES.keys())
    for k in keys:
        w = ITEM_TYPES[k]["weight"]
        if k == "bomb":
            w = int(w + level * 4)
        weights.append(w)
    total = sum(weights)
    r = random.uniform(0, total)
    acc = 0
    for k, w in zip(keys, weights):
        acc += w
        if r <= acc:
            return k
    return keys[-1]


class Item:
    __slots__ = ("kind", "x", "y", "vy", "spin", "id_main", "id_extra")

    def __init__(self, canvas, kind, x, y, vy):
        self.kind = kind
        self.x = x
        self.y = y
        self.vy = vy
        self.spin = random.uniform(0, 360)
        self.id_main = None
        self.id_extra = None
        self._draw(canvas)

    def _draw(self, canvas):
        spec = ITEM_TYPES[self.kind]
        c = spec["color"]
        r = spec["radius"]
        x, y = self.x, self.y
        shape = spec["shape"]
        if shape == "star":
            pts = []
            for i in range(10):
                ang = (i * 36 - 90) * 3.14159 / 180
                rad = r if i % 2 == 0 else r * 0.45
                import math
                pts += [x + math.cos(ang) * rad, y + math.sin(ang) * rad]
            self.id_main = canvas.create_polygon(pts, fill=c, outline="#fff8c8")
        elif shape == "diamond":
            pts = [x, y - r, x + r, y, x, y + r, x - r, y]
            self.id_main = canvas.create_polygon(pts, fill=c, outline="#e6ffff", width=2)
        elif shape == "bomb":
            self.id_main = canvas.create_oval(x - r, y - r, x + r, y + r, fill=c, outline="#2a0000", width=2)
            self.id_extra = canvas.create_line(x, y - r, x + 4, y - r - 7, fill="#ffaa44", width=3)
        elif shape == "ring":
            self.id_main = canvas.create_oval(x - r, y - r, x + r, y + r, outline=c, width=3)
            letter = "S" if self.kind == "shield" else "%"
            self.id_extra = canvas.create_text(x, y, text=letter, fill=c, font=("Consolas", 11, "bold"))

    def move(self, canvas, dt):
        dy = self.vy * dt
        self.y += dy
        canvas.move(self.id_main, 0, dy)
        if self.id_extra is not None:
            canvas.move(self.id_extra, 0, dy)

    def destroy(self, canvas):
        canvas.delete(self.id_main)
        if self.id_extra is not None:
            canvas.delete(self.id_extra)


class Particle:
    __slots__ = ("id", "vx", "vy", "life")

    def __init__(self, canvas, x, y, color):
        r = random.uniform(2, 4)
        self.id = canvas.create_oval(x - r, y - r, x + r, y + r, fill=color, outline="")
        ang = random.uniform(0, 6.283)
        spd = random.uniform(2, 6)
        import math
        self.vx = math.cos(ang) * spd
        self.vy = math.sin(ang) * spd - 1.5
        self.life = random.uniform(0.4, 0.9)

    def step(self, canvas, dt):
        self.vy += 0.25
        canvas.move(self.id, self.vx, self.vy)
        self.life -= dt / 60.0
        return self.life > 0


class Game:
    def __init__(self, root):
        self.root = root
        root.title("Cosmic Catcher")
        root.resizable(False, False)
        root.configure(bg="#05060d")

        self.canvas = tk.Canvas(root, width=W, height=H, bg="#05060d", highlightthickness=0)
        self.canvas.pack()

        # Starfield (parallax background)
        self.stars = []
        for _ in range(80):
            sx = random.uniform(0, W)
            sy = random.uniform(0, H)
            sz = random.choice([1, 1, 2, 3])
            sid = self.canvas.create_oval(sx, sy, sx + sz, sy + sz,
                                          fill="#ffffff" if sz == 3 else "#9aa4d6",
                                          outline="")
            self.stars.append([sid, sx, sy, sz, sz * 0.4 + 0.2])

        # Player ship
        cx = W / 2
        cy = H - 40
        self.px = cx
        self.py = cy
        self.player_id = self.canvas.create_polygon(
            cx - PLAYER_W / 2, cy + PLAYER_H / 2,
            cx,                cy - PLAYER_H / 2,
            cx + PLAYER_W / 2, cy + PLAYER_H / 2,
            cx + PLAYER_W / 4, cy + PLAYER_H / 4,
            cx - PLAYER_W / 4, cy + PLAYER_H / 4,
            fill="#7af6ff", outline="#ffffff", width=2,
        )
        self.thruster_id = self.canvas.create_polygon(
            cx - 8, cy + PLAYER_H / 2,
            cx,     cy + PLAYER_H / 2 + 14,
            cx + 8, cy + PLAYER_H / 2,
            fill="#ff8a3c", outline="",
        )

        # HUD
        self.hud_score   = self.canvas.create_text(14, 12, anchor="nw", fill="#ffd54a",
                                                   font=("Consolas", 14, "bold"), text="SCORE 0")
        self.hud_lives   = self.canvas.create_text(W - 14, 12, anchor="ne", fill="#ff7777",
                                                   font=("Consolas", 14, "bold"), text="LIVES 3")
        self.hud_level   = self.canvas.create_text(W / 2, 12, anchor="n", fill="#9aa4d6",
                                                   font=("Consolas", 12, "bold"), text="WAVE 1")
        self.hud_combo   = self.canvas.create_text(W / 2, 34, anchor="n", fill="#7af6ff",
                                                   font=("Consolas", 11), text="")
        self.hud_status  = self.canvas.create_text(W / 2, H / 2, anchor="c", fill="#ffffff",
                                                   font=("Consolas", 22, "bold"), text="")
        self.hud_help    = self.canvas.create_text(W / 2, H - 12, anchor="s", fill="#4d567a",
                                                   font=("Consolas", 9),
                                                   text="←/→ move   space brake   P pause   R restart   Esc quit")

        # Input
        self.keys = set()
        root.bind("<KeyPress>",   self._on_press)
        root.bind("<KeyRelease>", self._on_release)
        root.protocol("WM_DELETE_WINDOW", root.destroy)

        self.reset()
        self.last_t = time.perf_counter()
        self.root.after(FRAME_MS, self.tick)

    # -------- input --------
    def _on_press(self, e):
        k = e.keysym.lower()
        self.keys.add(k)
        if k == "escape":
            self.root.destroy()
        elif k == "p" and not self.game_over:
            self.paused = not self.paused
            self.canvas.itemconfigure(self.hud_status, text="PAUSED" if self.paused else "")
        elif k == "r" and self.game_over:
            self.reset()

    def _on_release(self, e):
        self.keys.discard(e.keysym.lower())

    # -------- state --------
    def reset(self):
        for it in getattr(self, "items", []):
            it.destroy(self.canvas)
        for p in getattr(self, "particles", []):
            self.canvas.delete(p.id)
        self.items = []
        self.particles = []
        self.score = 0
        self.lives = 3
        self.combo = 0
        self.best_combo = 0
        self.level = 1
        self.spawn_timer = 0.0
        self.spawn_interval = 38.0  # frames
        self.shield_until = 0.0
        self.slow_until = 0.0
        self.paused = False
        self.game_over = False
        self.canvas.itemconfigure(self.hud_status, text="")
        self._refresh_hud()

    def _refresh_hud(self):
        self.canvas.itemconfigure(self.hud_score, text=f"SCORE {self.score}")
        self.canvas.itemconfigure(self.hud_lives, text=f"LIVES {self.lives}")
        self.canvas.itemconfigure(self.hud_level, text=f"WAVE {self.level}")
        if self.combo >= 3:
            mult = self._combo_mult()
            self.canvas.itemconfigure(self.hud_combo, text=f"COMBO x{mult}  ({self.combo})")
        else:
            self.canvas.itemconfigure(self.hud_combo, text="")

    def _combo_mult(self):
        if self.combo >= 25: return 5
        if self.combo >= 15: return 4
        if self.combo >= 9:  return 3
        if self.combo >= 3:  return 2
        return 1

    # -------- main loop --------
    def tick(self):
        now = time.perf_counter()
        dt_real = now - self.last_t
        self.last_t = now
        # Convert to "frames at 60fps" so existing tuned values still feel right.
        dt = max(0.0, min(3.0, dt_real * 60.0))

        if not self.paused and not self.game_over:
            time_scale = 0.45 if now < self.slow_until else 1.0
            self._update(dt * time_scale, now)

        self.root.after(FRAME_MS, self.tick)

    def _update(self, dt, now):
        # --- starfield drift ---
        for s in self.stars:
            sid, sx, sy, sz, sp = s
            sy += sp * dt
            if sy > H:
                sy = 0
                sx = random.uniform(0, W)
                self.canvas.coords(sid, sx, sy, sx + sz, sy + sz)
            else:
                self.canvas.move(sid, 0, sp * dt)
            s[2] = sy

        # --- player movement ---
        speed = PLAYER_SPEED * (0.5 if "space" in self.keys else 1.0)
        if "left" in self.keys or "a" in self.keys:
            self.px -= speed * dt
        if "right" in self.keys or "d" in self.keys:
            self.px += speed * dt
        self.px = max(PLAYER_W / 2, min(W - PLAYER_W / 2, self.px))
        self._reposition_player()

        # thruster flicker
        flicker = random.uniform(0.6, 1.0)
        flame_color = "#ff8a3c" if flicker > 0.75 else "#ffce5c"
        self.canvas.itemconfigure(self.thruster_id, fill=flame_color)

        # --- spawn ---
        self.spawn_timer += dt
        if self.spawn_timer >= self.spawn_interval:
            self.spawn_timer = 0.0
            self._spawn_item()

        # --- items ---
        catch_y = self.py - PLAYER_H / 2
        keep = []
        for it in self.items:
            it.move(self.canvas, dt)
            if it.y > H + 20:
                # missed - star/gem breaks combo
                if it.kind in ("star", "gem"):
                    self.combo = 0
                it.destroy(self.canvas)
                continue
            # collision (AABB-ish vs ship triangle bbox)
            if (it.y >= catch_y and it.y <= self.py + PLAYER_H / 2
                    and abs(it.x - self.px) < PLAYER_W / 2 + 6):
                self._collect(it, now)
                continue
            keep.append(it)
        self.items = keep

        # --- particles ---
        pkeep = []
        for p in self.particles:
            if p.step(self.canvas, dt):
                pkeep.append(p)
            else:
                self.canvas.delete(p.id)
        self.particles = pkeep

        # --- level up ---
        target_level = 1 + self.score // 500
        if target_level > self.level:
            self.level = target_level
            self.spawn_interval = max(10.0, 38.0 - (self.level - 1) * 2.4)
            self._flash_status(f"WAVE {self.level}")

        self._refresh_hud()

    def _reposition_player(self):
        cx, cy = self.px, self.py
        self.canvas.coords(
            self.player_id,
            cx - PLAYER_W / 2, cy + PLAYER_H / 2,
            cx,                cy - PLAYER_H / 2,
            cx + PLAYER_W / 2, cy + PLAYER_H / 2,
            cx + PLAYER_W / 4, cy + PLAYER_H / 4,
            cx - PLAYER_W / 4, cy + PLAYER_H / 4,
        )
        self.canvas.coords(
            self.thruster_id,
            cx - 8, cy + PLAYER_H / 2,
            cx,     cy + PLAYER_H / 2 + 14,
            cx + 8, cy + PLAYER_H / 2,
        )

    def _spawn_item(self):
        kind = weighted_pick(self.level)
        x = random.uniform(20, W - 20)
        base_v = 2.4 + (self.level - 1) * 0.35
        vy = base_v + random.uniform(-0.4, 0.9)
        self.items.append(Item(self.canvas, kind, x, -20, vy))

    def _collect(self, it, now):
        spec = ITEM_TYPES[it.kind]
        if it.kind == "bomb":
            if now < self.shield_until:
                self._burst(it.x, it.y, "#6cff6c")
                self._flash_status("SHIELDED!")
            else:
                self.lives -= 1
                self.combo = 0
                self._burst(it.x, it.y, "#ff4a4a")
                self._flash_status("HIT!")
                if self.lives <= 0:
                    self._end_game()
        elif it.kind == "shield":
            self.shield_until = now + 2.0
            self._burst(it.x, it.y, "#6cff6c")
            self._flash_status("SHIELD UP")
        elif it.kind == "slow":
            self.slow_until = now + 3.0
            self._burst(it.x, it.y, "#c46cff")
            self._flash_status("SLOW-MO")
        else:
            self.combo += 1
            self.best_combo = max(self.best_combo, self.combo)
            mult = self._combo_mult()
            self.score += spec["score"] * mult
            self._burst(it.x, it.y, spec["color"])
        it.destroy(self.canvas)

    def _burst(self, x, y, color):
        for _ in range(14):
            self.particles.append(Particle(self.canvas, x, y, color))

    def _flash_status(self, msg):
        self.canvas.itemconfigure(self.hud_status, text=msg)
        self.root.after(550, lambda: (
            self.canvas.itemconfigure(self.hud_status, text="")
            if not self.paused and not self.game_over else None
        ))

    def _end_game(self):
        self.game_over = True
        self.canvas.itemconfigure(
            self.hud_status,
            text=f"GAME OVER\nscore {self.score}   best combo x{self.best_combo}\npress R to restart",
            justify="center",
        )


def main():
    root = tk.Tk()
    Game(root)
    root.mainloop()


if __name__ == "__main__":
    main()
