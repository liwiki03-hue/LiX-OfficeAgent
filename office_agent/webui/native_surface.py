import ctypes
from ctypes import wintypes

from PIL import Image


class Blend(ctypes.Structure):
    _fields_ = [('op', ctypes.c_ubyte), ('flags', ctypes.c_ubyte),
                ('alpha', ctypes.c_ubyte), ('format', ctypes.c_ubyte)]


class BitmapHeader(ctypes.Structure):
    _fields_ = [('size', wintypes.DWORD), ('width', wintypes.LONG), ('height', wintypes.LONG),
                ('planes', wintypes.WORD), ('bits', wintypes.WORD), ('compression', wintypes.DWORD),
                ('image_size', wintypes.DWORD), ('x', wintypes.LONG), ('y', wintypes.LONG),
                ('used', wintypes.DWORD), ('important', wintypes.DWORD)]


class PixelSurface:
    def __init__(self, form, *, clickthrough=False, noactivate=True):
        self.form = form
        self.hwnd = ctypes.c_void_p(int(form.Handle.ToInt64()))
        self.user = ctypes.WinDLL('user32', use_last_error=True)
        self.gdi = ctypes.WinDLL('gdi32', use_last_error=True)
        pointer = ctypes.c_void_p
        self.user.GetWindowLongW.argtypes = [pointer, ctypes.c_int]
        self.user.SetWindowLongW.argtypes = [pointer, ctypes.c_int, ctypes.c_long]
        style = self.user.GetWindowLongW(self.hwnd, -20)
        self.user.SetWindowLongW(self.hwnd, -20, style | 0x00080000 | (0x08000000 if noactivate else 0) | (0x20 if clickthrough else 0))
        self.gdi.CreateCompatibleDC.argtypes = [pointer]
        self.gdi.CreateCompatibleDC.restype = pointer
        self.gdi.CreateDIBSection.argtypes = [pointer, ctypes.POINTER(BitmapHeader), wintypes.UINT, ctypes.POINTER(pointer), pointer, wintypes.DWORD]
        self.gdi.CreateDIBSection.restype = pointer
        self.gdi.SelectObject.argtypes = [pointer, pointer]
        self.gdi.SelectObject.restype = pointer
        self.gdi.DeleteObject.argtypes = [pointer]
        self.gdi.DeleteDC.argtypes = [pointer]
        self.user.UpdateLayeredWindow.argtypes = [pointer, pointer, ctypes.POINTER(wintypes.POINT), ctypes.POINTER(wintypes.SIZE), pointer, ctypes.POINTER(wintypes.POINT), wintypes.DWORD, ctypes.POINTER(Blend), wintypes.DWORD]
        self.draws = 0

    def draw(self, image):
        if image.mode != 'RGBA' or image.width > 1024 or image.height > 1024:
            raise ValueError('Invalid native surface')
        premultiplied = Image.frombytes('RGBA', image.size, image.convert('RGBa').tobytes()).tobytes('raw', 'BGRA')
        dc = self.gdi.CreateCompatibleDC(None)
        bitmap = previous = None
        try:
            header = BitmapHeader(size=ctypes.sizeof(BitmapHeader), width=image.width, height=-image.height, planes=1, bits=32)
            pixels = ctypes.c_void_p()
            bitmap = self.gdi.CreateDIBSection(dc, ctypes.byref(header), 0, ctypes.byref(pixels), None, 0)
            if not dc or not bitmap or not pixels:
                raise ctypes.WinError(ctypes.get_last_error())
            ctypes.memmove(pixels, premultiplied, len(premultiplied))
            previous = self.gdi.SelectObject(dc, bitmap)
            destination = wintypes.POINT(self.form.Left, self.form.Top)
            size = wintypes.SIZE(image.width, image.height)
            source = wintypes.POINT(0, 0)
            blend = Blend(0, 0, 255, 1)
            if not self.user.UpdateLayeredWindow(self.hwnd, None, ctypes.byref(destination), ctypes.byref(size), dc, ctypes.byref(source), 0, ctypes.byref(blend), 2):
                raise ctypes.WinError(ctypes.get_last_error())
            self.draws += 1
        finally:
            if previous:
                self.gdi.SelectObject(dc, previous)
            if bitmap:
                self.gdi.DeleteObject(bitmap)
            if dc:
                self.gdi.DeleteDC(dc)
