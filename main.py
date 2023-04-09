import win32gui
import win32con

from loguru import logger
from pynput import keyboard

from key_map import KEY_MAP

class MultiControl:
    def __init__(self):
        self.game_window_name = "Toontown Rewritten"
        self.game_window_class = "WinGraphicsWindow0"
        self.game_handles = self.get_game_handles()

        # Create a set to keep track of currently pressed keys
        self.pressed_keys = set()


    def get_game_handle(self):
        """Gets handle to game window."""
        logger.debug("get_game_handle called")
        return win32gui.FindWindow(self.game_window_class, self.game_window_name)


    def get_game_handles(self):
        """Gets handles to all game windows."""
        game_handles = []

        def enumHandler(hwnd, lParam):
            window_class = win32gui.GetClassName(hwnd)
            window_title = win32gui.GetWindowText(hwnd)
            
            if self.game_window_name in window_title and window_class == self.game_window_class:
                logger.info(f"Found game window: {window_title} (handle: {hwnd})")
                game_handles.append(hwnd)

        win32gui.EnumWindows(enumHandler, None)
        return game_handles


    def press_key(self, handle, hotkey):
        """Gets key map for hotkey, presses the hotkey, and waits a randomized fraction of a second."""
        keycode = KEY_MAP[hotkey]
        logger.debug(f"Pressing key: {hotkey} (keycode: {keycode})")
        win32gui.SendMessage(handle, win32con.WM_KEYDOWN, keycode, 0)


    def release_key(self, handle, hotkey):
        keycode = KEY_MAP[hotkey]
        logger.debug(f"Releasing key: {hotkey} (keycode: {keycode})")
        win32gui.SendMessage(handle, win32con.WM_KEYUP, keycode, 0)


    def on_press(self, key):
        try:
            key_pressed = key.char  # single-char keys
        except:
            key_pressed = key.name  # other keys

        # Skip if the key is already pressed
        if key_pressed in self.pressed_keys:
            return

        self.pressed_keys.add(key_pressed)

        active_window_handle = win32gui.GetForegroundWindow()
        active_window_title = win32gui.GetWindowText(active_window_handle)

        if active_window_title == self.game_window_name:
            for handle in self.game_handles:
                if handle != active_window_handle:
                    self.press_key(handle, key_pressed)


    def on_release(self, key):
        try:
            key_released = key.char  # single-char keys
        except:
            key_released = key.name  # other keys

        # Remove the key from the pressed keys set
        if key_released in self.pressed_keys:
            self.pressed_keys.remove(key_released)

        active_window_handle = win32gui.GetForegroundWindow()
        active_window_title = win32gui.GetWindowText(active_window_handle)

        if active_window_title == self.game_window_name:
            for handle in self.game_handles:
                if handle != active_window_handle:
                    self.release_key(handle, key_released)


    def start_key_listener(self):
        logger.info("Starting key listener")
        self.listener = keyboard.Listener(on_press=self.on_press, on_release=self.on_release)
        self.listener.start()
        self.listener.join()


def main():
    controller = MultiControl()
    controller.start_key_listener()



if __name__ == "__main__":
    main()
