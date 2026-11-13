import math
import time


class Transition:
    def __init__(self, value=0.0):
        self.start = self.target = float(value)
        self.started = self.duration = 0.0

    def value(self, now):
        if self.duration <= 0:
            return self.target
        progress = max(0.0, min(1.0, (now - self.started) / self.duration))
        return self.start + (self.target - self.start) * (1 - (1 - progress) ** 3)

    def move(self, target, now, duration):
        current = self.value(now)
        self.start, self.target, self.started = current, float(target), now
        self.duration = duration * abs(self.target - current)

    def running(self, now):
        return self.duration > 0 and now < self.started + self.duration


class OrbMotion:
    def __init__(self, clock=time.monotonic):
        self.clock = clock
        self.last = clock()
        self.elapsed = self.angle = 0.0
        self.status = 'ready'
        self.changed = self.last - 1
        self.visible = False
        self.reduced = self.dragging = self.hovered = self.focused = self.drop = self.expanded = False
        self.emphasis = Transition()
        self.menu = Transition()

    def update(self, *, visible=None, reduced=None, dragging=None, hovered=None, focused=None, drop=None, expanded=None, status=None):
        now = self.clock()
        self._advance(now)
        if status is not None and status != self.status:
            self.status, self.changed = status, now
        for name, value in (('visible', visible), ('reduced', reduced), ('dragging', dragging), ('hovered', hovered), ('focused', focused), ('drop', drop)):
            if value is not None:
                setattr(self, name, bool(value))
        target = float(self.hovered or self.focused or self.drop)
        if target != self.emphasis.target or self.reduced:
            self.emphasis.move(target, now, 0 if self.reduced else .16)
        if expanded is not None:
            self.expanded = bool(expanded)
        if float(self.expanded) != self.menu.target or self.reduced or not self.visible:
            self.menu.move(float(self.expanded), now, 0 if self.reduced or not self.visible else .24)

    def _advance(self, now):
        delta = max(0.0, now - self.last)
        if self.visible and not self.reduced and not self.dragging:
            self.elapsed += delta
            speed = {'ready': 1 / 24, 'opening': 1 / 6, 'busy': 1 / 8}.get(self.status, 0)
            self.angle = (self.angle + delta * speed * math.tau) % math.tau
        self.last = now

    def interval(self):
        now = self.clock()
        if not self.visible or self.reduced or self.dragging:
            return 0
        if self.menu.running(now) or self.emphasis.running(now):
            return 30
        if self.expanded:
            return 0
        if self.status in {'complete', 'error', 'stopped', 'uncertain'}:
            return 30 if now - self.changed < .6 else 0
        if self.status in {'opening', 'busy'} or self.drop or self.hovered or self.focused:
            return 30 if self.status != 'confirm' else 0
        return 83 if self.status == 'ready' else 0

    def frame(self):
        now = self.clock()
        self._advance(now)
        moving = not self.reduced and not self.dragging and self.visible
        wave = math.sin(self.elapsed * math.tau / 4.8) if moving and self.status in {'ready', 'opening', 'busy'} else 0
        age = now - self.changed
        feedback = math.sin(math.pi * age / .6) if moving and self.status in {'complete', 'error', 'stopped', 'uncertain'} and 0 <= age < .6 else 0
        return {'phase': round(self.angle / math.tau * 360) % 360 if not self.reduced else 0,
                'pulse': round(wave * 5 + feedback * 5), 'emphasis': round(self.emphasis.value(now) * 8),
                'menu': self.menu.value(now), 'menu_running': self.menu.running(now)}
