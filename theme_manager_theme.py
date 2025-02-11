from PyQt5.QtWidgets import QApplication
from qt_material import apply_stylesheet, list_themes

class ThemeManager:
    def __init__(self, app):
        self.app = app
        self.current_theme = "light_blue.xml"
        self.themes = list_themes()

    def set_theme(self, theme):
        if theme in self.themes:
            self.current_theme = theme
            apply_stylesheet(self.app, theme=theme)
            self._adjust_button_style()
        else:
            print(f"Theme {theme} not found. Using default theme.")
            self.set_default_theme()

    def set_default_theme(self):
        self.current_theme = "light_blue.xml"
        apply_stylesheet(self.app, theme="light_blue.xml")
        self._adjust_button_style()

    def toggle_theme(self):
        current_index = self.themes.index(self.current_theme)
        next_index = (current_index + 1) % len(self.themes)
        self.set_theme(self.themes[next_index])

    def get_available_themes(self):
        return self.themes

    def _adjust_button_style(self):
        if "light" in self.current_theme:
            self.app.setStyleSheet(self.app.styleSheet() + """
                QPushButton:disabled {
                    color: #808080;
                }
                QPushButton[enabled="false"] {
                    color: #808080;
                }
            """)
        else:
            self.app.setStyleSheet(self.app.styleSheet() + """
                QPushButton:disabled {
                    color: #808080;
                }
                QPushButton[enabled="false"] {
                    color: #808080;
                }
            """)
