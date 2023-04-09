import ctypes
import sys
import time
import win32gui
import win32con
import win32api

from loguru import logger
from pynput import keyboard
from pynput.mouse import Listener as MouseListener, Button as MouseButton


from key_map import KEY_MAP


class MultiControl:
    def __init__(self):
        self.game_window_name = "Toontown Rewritten"
        self.game_window_class = "WinGraphicsWindow0"
        self.game_handles = self.get_game_handles()

        # Create a set to keep track of currently pressed keys
        self.pressed_keys = set()

        self.stopped = True

###############
# GAME WINDOW #
###############

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


    def get_window_border_and_titlebar_dimensions(self):
        border_width = win32api.GetSystemMetrics(win32con.SM_CXSIZEFRAME)
        border_height = win32api.GetSystemMetrics(win32con.SM_CYSIZEFRAME)
        titlebar_height = win32api.GetSystemMetrics(win32con.SM_CYCAPTION)
        return border_width, border_height, titlebar_height


    def bring_game_windows_to_foreground(self, game_handles):
        for handle in game_handles:
            logger.info(f"Bringing game window (handle: {handle}) to foreground")

            # Store the original foreground window
            original_foreground_window = win32gui.GetForegroundWindow()

            # Set the game window as the foreground window temporarily
            win32gui.SetForegroundWindow(handle)
            time.sleep(0.1)

            # Simulate Alt+Tab to bring the game window to the foreground
            ctypes.windll.user32.keybd_event(win32con.VK_MENU, 0, 0, 0)
            ctypes.windll.user32.keybd_event(win32con.VK_TAB, 0, 0, 0)
            ctypes.windll.user32.keybd_event(win32con.VK_TAB, 0, win32con.KEYEVENTF_KEYUP, 0)
            ctypes.windll.user32.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)

            # Restore the original foreground window
            win32gui.SetForegroundWindow(original_foreground_window)
            time.sleep(0.5)  # Add a small delay between bringing each window to the foreground


    def press_key(self, handle, hotkey):
        """Gets key map for hotkey, presses the hotkey, and waits a randomized fraction of a second."""
        try:
            keycode = KEY_MAP[hotkey]
        except KeyError:
            logger.error(f"Invalid hotkey: {hotkey}")
            return
        logger.debug(f"Pressing key: {hotkey} (keycode: {keycode})")
        win32gui.SendMessage(handle, win32con.WM_KEYDOWN, keycode, 0)


    def release_key(self, handle, hotkey):
        keycode = KEY_MAP[hotkey]
        logger.debug(f"Releasing key: {hotkey} (keycode: {keycode})")
        win32gui.SendMessage(handle, win32con.WM_KEYUP, keycode, 0)

################
# KEY LISTENER #
################

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
            if key_pressed == PAUSE_KEY:
                if self.stopped:
                    logger.info("Resuming mirroring")
                    self.stopped = False
                else:
                    logger.info("Pausing mirroring")
                    self.stopped = True
            # Mirror the keypress if not paused.
            if not self.stopped:
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

##################
# MOUSE LISTENER #
##################

    def on_click(self, x, y, button, pressed):
        if self.stopped:
            return
        active_window_handle = win32gui.GetForegroundWindow()
        active_window_title = win32gui.GetWindowText(active_window_handle)

        if active_window_title == self.game_window_name and active_window_handle in self.game_handles:
            active_window_rect = win32gui.GetWindowRect(active_window_handle)
            relative_x = x - active_window_rect[0]
            relative_y = y - active_window_rect[1]

            # Adjust the click coordinates for the window border and title bar dimensions
            border_width, border_height, titlebar_height = self.get_window_border_and_titlebar_dimensions()
            relative_x -= border_width
            relative_y -= (border_height + titlebar_height)

            # Temporarily remove the active window handle from the game_handles list
            temp_game_handles = self.game_handles.copy()
            temp_game_handles.remove(active_window_handle)

            for handle in temp_game_handles:
                if pressed:
                    self.press_mouse_button(handle, button, relative_x, relative_y)
                else:
                    self.release_mouse_button(handle, button, relative_x, relative_y)


    def press_mouse_button(self, handle, button, x, y):
        lParam = win32api.MAKELONG(x, y)

        if button == MouseButton.left:
            win32gui.SendMessage(handle, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lParam)
        elif button == MouseButton.right:
            win32gui.SendMessage(handle, win32con.WM_RBUTTONDOWN, win32con.MK_RBUTTON, lParam)


    def release_mouse_button(self, handle, button, x, y):
        lParam = win32api.MAKELONG(x, y)

        if button == MouseButton.left:
            win32gui.SendMessage(handle, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, lParam)
        elif button == MouseButton.right:
            win32gui.SendMessage(handle, win32con.WM_RBUTTONUP, win32con.MK_RBUTTON, lParam)


    def start_listeners(self, key_listener=True, mouse_listener=True):
        logger.info("Starting listeners")
        self.stopped = False
        if key_listener:
            logger.info("Starting key listener")
            self.key_listener = keyboard.Listener(on_press=self.on_press, on_release=self.on_release)
            self.key_listener.start()
        if mouse_listener:
            logger.info("Starting mouse listener")
            self.mouse_listener = MouseListener(on_click=self.on_click)
            self.mouse_listener.start()


    def stop_listeners(self):
        logger.info("Stopping listeners")
        self.stopped = True
        if self.key_listener:
            logger.info("Stopping key listener")
            self.key_listener.stop()
        if self.mouse_listener:
            logger.info("Stopping mouse listener")
            self.mouse_listener.stop()


PAUSE_KEY = "`"
USE_KEY_LISTENER = True
USE_MOUSE_LISTENER = True

def main():
    # Set log level to INFO. Change to DEBUG if needed.
    logger.remove()
    logger.add(sys.stderr, level="INFO")

    # Start the key and mouse listeners for multiboxing.
    controller = MultiControl()
    controller.bring_game_windows_to_foreground(controller.game_handles)
    controller.start_listeners(key_listener=USE_KEY_LISTENER, mouse_listener=USE_MOUSE_LISTENER)

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        logger.info("Exiting...")
        controller.stop_listeners()


if __name__ == "__main__":
    main()
