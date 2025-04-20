# -*- coding: utf-8 -*-
from PyQt6.QtGui import QPalette, QColor
from PyQt6.QtWidgets import QApplication

class ThemeManager:
    """Manage application themes and colors"""

    # Theme colors (Keep the THEMES dictionary exactly as in your original script)
    THEMES = {
        "Light": {
            "bg_primary": "#FFFFFF",
            "bg_secondary": "#F5F5F5",
            "text_primary": "#333333",
            "text_secondary": "#666666",
            "accent": "#1976D2",
            "accent_hover": "#1565C0",
            "success": "#4CAF50",
            "warning": "#FF9800",
            "error": "#F44336",
            "button_bg": "#E0E0E0",
            "button_text": "#333333",
            "button_hover": "#BDBDBD",
            "border": "#DDDDDD"
        },
        "Dark": {
            "bg_primary": "#2D2D2D",
            "bg_secondary": "#424242",
            "text_primary": "#E0E0E0",
            "text_secondary": "#BDBDBD",
            "accent": "#2196F3",
            "accent_hover": "#1E88E5",
            "success": "#66BB6A",
            "warning": "#FFA726",
            "error": "#EF5350",
            "button_bg": "#5A5A5A",
            "button_text": "#E0E0E0",
            "button_hover": "#787878",
            "border": "#5D5D5D"
        },
        "Blue": {
            "bg_primary": "#E3F2FD",
            "bg_secondary": "#BBDEFB",
            "text_primary": "#1A237E",
            "text_secondary": "#3949AB",
            "accent": "#1976D2",
            "accent_hover": "#1565C0",
            "success": "#4CAF50",
            "warning": "#FF9800",
            "error": "#F44336",
            "button_bg": "#64B5F6",
            "button_text": "#FFFFFF",
            "button_hover": "#42A5F5",
            "border": "#90CAF9"
        }
    }

    @staticmethod
    def apply_theme(app: QApplication, theme_name: str = "Light"):
        """Apply a theme to the application"""

        if theme_name not in ThemeManager.THEMES:
            theme_name = "Light"

        theme = ThemeManager.THEMES[theme_name]

        # Create a palette
        palette = QPalette()

        # Set global colors
        palette.setColor(QPalette.ColorRole.Window, QColor(theme["bg_primary"]))
        palette.setColor(QPalette.ColorRole.WindowText, QColor(theme["text_primary"]))
        palette.setColor(QPalette.ColorRole.Base, QColor(theme["bg_secondary"]))
        palette.setColor(QPalette.ColorRole.AlternateBase, QColor(theme["bg_primary"]))
        palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(theme["bg_secondary"]))
        palette.setColor(QPalette.ColorRole.ToolTipText, QColor(theme["text_primary"]))
        palette.setColor(QPalette.ColorRole.Text, QColor(theme["text_primary"]))
        palette.setColor(QPalette.ColorRole.Button, QColor(theme["button_bg"]))
        palette.setColor(QPalette.ColorRole.ButtonText, QColor(theme["button_text"]))
        palette.setColor(QPalette.ColorRole.BrightText, QColor(theme["text_primary"])) # Often same as WindowText or white for dark themes
        palette.setColor(QPalette.ColorRole.Link, QColor(theme["accent"]))
        palette.setColor(QPalette.ColorRole.Highlight, QColor(theme["accent"]))
        palette.setColor(QPalette.ColorRole.HighlightedText, QColor(theme["bg_primary"])) # Or white/appropriate contrast

        # Set disabled colors (important for usability)
        palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.WindowText, QColor("#A0A0A0"))
        palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Text, QColor("#A0A0A0"))
        palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.ButtonText, QColor("#A0A0A0"))
        palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Button, QColor("#D0D0D0")) # Example disabled button bg

        # Apply the palette
        app.setPalette(palette)

        return theme