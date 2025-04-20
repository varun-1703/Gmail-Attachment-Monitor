# -*- coding: utf-8 -*-
"""
Custom PyQt6 Widgets for the Gmail Attachment Monitor application.

Provides themed widgets inheriting from standard PyQt6 classes
with enhanced styling and specific functionalities like pulsing progress bars.
"""

from typing import Optional, Dict, Any

from PyQt6.QtWidgets import (
    QProgressBar, QPushButton, QGroupBox, QTreeWidget, QTextEdit, QLineEdit,
    QSpinBox, QComboBox, QFrame, QLabel, QVBoxLayout, QHBoxLayout, QWidget,
    QAbstractItemView, QHeaderView, QListView # <-- Added QListView here
)
from PyQt6.QtCore import Qt, QTimer, QSize
from PyQt6.QtGui import QFont, QIcon

# Type alias for theme dictionary for clarity
ThemeDict = Dict[str, str]

# --- Progress Bars ---

class StyledProgressBar(QProgressBar):
    """Custom styled progress bar for displaying fixed progress."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setTextVisible(False)
        self.setFixedHeight(8)
        self.setMaximum(100)
        self.setValue(0)
        # Initial style (will be overridden by setTheme)
        self.setStyleSheet("""
            QProgressBar { border: none; border-radius: 4px; background-color: #E0E0E0; }
            QProgressBar::chunk { background-color: #1976D2; border-radius: 4px; }
        """)

    def setTheme(self, theme: ThemeDict):
        """Applies theme colors to the progress bar."""
        self.setStyleSheet(f"""
            QProgressBar {{
                border: none;
                border-radius: 4px;
                background-color: {theme.get('bg_secondary', '#E0E0E0')};
                height: 8px; /* Ensure height is maintained */
                text-align: center; /* Center text if ever made visible */
            }}
            QProgressBar::chunk {{
                background-color: {theme.get('accent', '#1976D2')};
                border-radius: 4px;
            }}
        """)


class PulseProgressBar(QProgressBar):
    """Progress bar that pulses visually when monitoring is active."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setTextVisible(False)
        self.setFixedHeight(4) # Make it slimmer for status bar use
        self.setMaximum(100)
        self.setValue(0)

        self.pulse_timer = QTimer(self)
        self.pulse_timer.timeout.connect(self._pulse_animation)
        self._pulse_value = 0
        self._pulse_direction = 1

        # Initial style (will be overridden by setTheme)
        self.setStyleSheet("""
            QProgressBar { border: none; border-radius: 2px; background-color: #E0E0E0; }
            QProgressBar::chunk { background-color: #1976D2; border-radius: 2px; }
        """)

    def start_pulse(self):
        """Starts the pulsing animation."""
        if not self.pulse_timer.isActive():
            self._pulse_value = 0
            self._pulse_direction = 1
            self.setValue(0)
            self.pulse_timer.start(40) # Adjust interval for pulse speed

    def stop_pulse(self):
        """Stops the pulsing animation and resets the bar."""
        if self.pulse_timer.isActive():
            self.pulse_timer.stop()
        self.setValue(0) # Reset to empty

    def _pulse_animation(self):
        """Internal method to update the pulse effect."""
        increment = 4 # Adjust speed/visibility of pulse
        self._pulse_value += self._pulse_direction * increment

        # Reverse direction at boundaries
        if self._pulse_value >= 100:
            self._pulse_value = 100
            self._pulse_direction = -1
        elif self._pulse_value <= 0:
            self._pulse_value = 0
            self._pulse_direction = 1

        self.setValue(self._pulse_value)

    def setTheme(self, theme: ThemeDict):
        """Applies theme colors to the pulsing progress bar."""
        self.setStyleSheet(f"""
            QProgressBar {{
                border: none;
                border-radius: 2px;
                background-color: {theme.get('bg_secondary', '#E0E0E0')};
                height: 4px; /* Ensure height is maintained */
            }}
            QProgressBar::chunk {{
                background-color: {theme.get('accent', '#1976D2')};
                border-radius: 2px;
            }}
        """)


# --- Buttons ---

class StyledButton(QPushButton):
    """Custom styled button with primary (accent) and secondary options."""

    def __init__(self,
                 text: str,
                 parent: Optional[QWidget] = None,
                 icon: Optional[QIcon] = None,
                 primary: bool = False):
        """
        Initializes the styled button.

        Args:
            text: Text displayed on the button.
            parent: Parent widget.
            icon: Optional QIcon for the button.
            primary: If True, use accent colors; otherwise, use standard button colors.
        """
        # Call the parent __init__ WITHOUT the 'primary' argument
        super().__init__(text, parent)

        # Store the primary flag as an instance variable
        self.is_primary = primary # Use a different name to avoid potential clashes

        if icon:
            self.setIcon(icon)
            # Consider setting icon size if needed: self.setIconSize(QSize(16, 16))

        self.setCursor(Qt.CursorShape.PointingHandCursor)
        # Theme will be applied later by the main window calling setTheme

    def setTheme(self, theme: ThemeDict):
        """Sets the visual theme for the button based on its type (primary/secondary)."""
        padding_vertical = 8
        padding_horizontal = 16
        border_radius = 4

        if self.is_primary:
            # Primary Button Style (Accent Color)
            self.setStyleSheet(f"""
                QPushButton {{
                    background-color: {theme.get('accent', '#1976D2')};
                    color: white; /* High contrast for accent */
                    border: none;
                    padding: {padding_vertical}px {padding_horizontal}px;
                    border-radius: {border_radius}px;
                    font-weight: bold;
                }}
                QPushButton:hover {{
                    background-color: {theme.get('accent_hover', '#1565C0')};
                }}
                QPushButton:pressed {{
                    background-color: {theme.get('accent', '#1976D2')};
                    /* Slight inset effect */
                    padding-top: {padding_vertical + 1}px;
                    padding-bottom: {padding_vertical - 1}px;
                }}
                QPushButton:disabled {{
                    background-color: {theme.get('button_hover', '#BDBDBD')}; /* Use a theme color */
                    color: {theme.get('text_secondary', '#757575')};
                }}
            """)
        else:
            # Secondary Button Style (Standard Theme Button Colors)
            self.setStyleSheet(f"""
                QPushButton {{
                    background-color: {theme.get('button_bg', '#E0E0E0')};
                    color: {theme.get('button_text', '#333333')};
                    border: 1px solid {theme.get('border', '#DDDDDD')}; /* Define border */
                    padding: {padding_vertical - 1}px {padding_horizontal}px; /* Adjust padding for border */
                    border-radius: {border_radius}px;
                    font-weight: normal; /* Normal weight for secondary */
                }}
                QPushButton:hover {{
                    background-color: {theme.get('button_hover', '#BDBDBD')};
                    border-color: {theme.get('button_hover', '#BDBDBD')}; /* Match border */
                }}
                QPushButton:pressed {{
                    background-color: {theme.get('button_bg', '#E0E0E0')};
                    /* Slight inset effect */
                    padding-top: {padding_vertical}px;
                    padding-bottom: {padding_vertical - 2}px;
                }}
                QPushButton:disabled {{
                    background-color: {theme.get('bg_secondary', '#F5F5F5')}; /* Lighter disabled bg */
                    color: {theme.get('text_secondary', '#AAAAAA')};
                    border-color: {theme.get('border', '#DDDDDD')};
                }}
            """)


# --- Containers ---

class StyledGroupBox(QGroupBox):
    """Custom styled group box with themed border and title."""

    def __init__(self, title: str, parent: Optional[QWidget] = None):
        super().__init__(title, parent)
        # Initial style (will be overridden by setTheme)
        self.setStyleSheet("""
            QGroupBox {
                border: 1px solid #DDDDDD; border-radius: 6px;
                margin-top: 10px; padding: 15px; font-weight: bold;
            }
            QGroupBox::title {
                subcontrol-origin: margin; subcontrol-position: top left;
                padding: 0 5px; color: #1976D2; left: 10px; /* Position title */
            }
        """)

    def setTheme(self, theme: ThemeDict):
        """Applies theme colors to the group box."""
        self.setStyleSheet(f"""
            QGroupBox {{
                border: 1px solid {theme.get('border', '#DDDDDD')};
                border-radius: 6px;
                margin-top: 10px; /* Space for the title */
                font-weight: bold;
                padding: 15px; /* Padding inside the box */
                padding-top: 20px; /* Extra padding at top for title */
                background-color: {theme.get('bg_primary', '#FFFFFF')}; /* Set background */
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 5px; /* Padding around the title text */
                left: 10px; /* Offset from left edge */
                color: {theme.get('accent', '#1976D2')}; /* Use accent color for title */
                background-color: {theme.get('bg_primary', '#FFFFFF')}; /* Match background */
            }}
        """)


# --- Data Display ---

class StyledTreeWidget(QTreeWidget):
    """Custom styled tree widget for displaying hierarchical data."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setAlternatingRowColors(True)
        self.setRootIsDecorated(False) # Cleaner look for flat lists
        self.setUniformRowHeights(True)
        self.setAnimated(True) # Subtle animation on expand/collapse
        self.setSortingEnabled(True)
        self.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        # Initial style (will be overridden by setTheme)
        self.setStyleSheet("QTreeWidget { border: 1px solid #DDDDDD; border-radius: 4px; }")

    def setTheme(self, theme: ThemeDict):
        """Applies theme colors to the tree widget and its header."""
        # Determine alternating row color based on primary/secondary background contrast
        # This provides subtle striping effect
        alt_row_color = theme.get('bg_secondary', '#F5F5F5')
        if theme.get('bg_primary') == '#FFFFFF': # Light theme default
            alt_row_color = '#F9F9F9' # Even subtler alternation for light theme
        elif theme.get('bg_primary') == '#2D2D2D': # Dark theme default
             alt_row_color = '#3A3A3A' # Subtler alternation for dark theme

        self.setStyleSheet(f"""
            QTreeWidget {{
                border: 1px solid {theme.get('border', '#DDDDDD')};
                border-radius: 4px;
                background-color: {theme.get('bg_primary', '#FFFFFF')};
                color: {theme.get('text_primary', '#333333')};
                alternate-background-color: {alt_row_color}; /* Use calculated alt color */
            }}
            QTreeWidget::item {{
                padding: 6px 5px; /* Adjust vertical/horizontal padding */
                /* border-bottom: 1px solid {theme.get('border', '#DDDDDD')}; */ /* Optional: item border */
            }}
            QTreeWidget::item:selected {{
                background-color: {theme.get('accent', '#1976D2')};
                color: white; /* High contrast selection text */
            }}
            QTreeWidget::item:hover:!selected {{
                background-color: {theme.get('button_hover', '#BDBDBD')}30; /* Use alpha transparency for subtle hover */
                color: {theme.get('text_primary', '#333333')}; /* Ensure text remains readable */
            }}
            QHeaderView::section {{
                background-color: {theme.get('bg_secondary', '#F5F5F5')};
                color: {theme.get('text_primary', '#333333')};
                padding: 5px;
                border: none; /* Cleaner header */
                border-bottom: 1px solid {theme.get('border', '#DDDDDD')};
                /* border-right: 1px solid {theme.get('border', '#DDDDDD')}; */ /* Optional right border */
                font-weight: bold;
                height: 30px; /* Ensure consistent header height */
            }}
            QHeaderView {{
                /* Needed on some platforms to prevent white background artifacts */
                background-color: {theme.get('bg_secondary', '#F5F5F5')};
            }}
        """)
        # Set header properties after setting stylesheet
        header = self.header()
        # Don't stretch last section by default, let main window configure it
        # header.setStretchLastSection(True)
        header.setSectionsClickable(True)


# --- Input Fields ---

class StyledTextEdit(QTextEdit):
    """Custom styled text edit area."""

    def __init__(self, parent: Optional[QWidget] = None, read_only: bool = True):
        super().__init__(parent)
        self.setReadOnly(read_only)
        # Initial style (will be overridden by setTheme)
        self.setStyleSheet("QTextEdit { border: 1px solid #DDDDDD; border-radius: 4px; padding: 8px; }")

    def setTheme(self, theme: ThemeDict):
        """Applies theme colors to the text edit."""
        self.setStyleSheet(f"""
            QTextEdit {{
                border: 1px solid {theme.get('border', '#DDDDDD')};
                border-radius: 4px;
                background-color: {theme.get('bg_secondary', '#F5F5F5')};
                color: {theme.get('text_primary', '#333333')};
                padding: 8px;
                selection-background-color: {theme.get('accent', '#1976D2')}; /* Text selection color */
                selection-color: white;
            }}
        """)


class StyledLineEdit(QLineEdit):
    """Custom styled single-line text input field."""

    def __init__(self, text: str = "", parent: Optional[QWidget] = None):
        super().__init__(text, parent)
        # Initial style (will be overridden by setTheme)
        self.setStyleSheet("QLineEdit { border: 1px solid #DDDDDD; border-radius: 4px; padding: 8px; }")

    def setTheme(self, theme: ThemeDict):
        """Applies theme colors to the line edit."""
        self.setStyleSheet(f"""
            QLineEdit {{
                border: 1px solid {theme.get('border', '#DDDDDD')};
                border-radius: 4px;
                padding: 8px;
                background-color: {theme.get('bg_secondary', '#F5F5F5')};
                color: {theme.get('text_primary', '#333333')};
                selection-background-color: {theme.get('accent', '#1976D2')};
                selection-color: white;
            }}
            QLineEdit:focus {{
                border: 1px solid {theme.get('accent', '#1976D2')}; /* Highlight border on focus */
                /* Optional: Slightly different background on focus */
                /* background-color: {theme.get('bg_primary', '#FFFFFF')}; */
            }}
            QLineEdit:read-only {{
                background-color: {theme.get('button_bg', '#E0E0E0')}; /* Different bg for read-only */
            }}
            QLineEdit:disabled {{
                background-color: {theme.get('button_hover', '#BDBDBD')}80; /* Semi-transparent disabled look */
                color: {theme.get('text_secondary', '#AAAAAA')};
            }}
        """)


class StyledSpinBox(QSpinBox):
    """Custom styled spin box for numerical input."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        # Initial style (will be overridden by setTheme)
        self.setStyleSheet("QSpinBox { border: 1px solid #DDDDDD; border-radius: 4px; padding: 8px; }")

    def setTheme(self, theme: ThemeDict):
        """Applies theme colors to the spin box."""
        self.setStyleSheet(f"""
            QSpinBox {{
                border: 1px solid {theme.get('border', '#DDDDDD')};
                border-radius: 4px;
                padding: 8px; /* Main padding */
                padding-right: 25px; /* Extra padding to prevent text overlap with buttons */
                background-color: {theme.get('bg_secondary', '#F5F5F5')};
                color: {theme.get('text_primary', '#333333')};
                selection-background-color: {theme.get('accent', '#1976D2')};
                selection-color: white;
                min-height: 20px; /* Ensure minimum height */
            }}
            QSpinBox:focus {{
                 border: 1px solid {theme.get('accent', '#1976D2')};
            }}
            QSpinBox::up-button {{
                subcontrol-origin: border; /* Position relative to border */
                subcontrol-position: top right; /* Position at the top right */
                width: 22px; /* Width of the button area */
                height: 50%; /* Half the height */
                border-left: 1px solid {theme.get('border', '#DDDDDD')};
                border-bottom: 1px solid {theme.get('border', '#DDDDDD')};
                border-top-right-radius: 3px; /* Match main radius */
                background-color: {theme.get('button_bg', '#E0E0E0')};
            }}
            QSpinBox::up-button:hover {{
                background-color: {theme.get('button_hover', '#BDBDBD')};
            }}
            QSpinBox::up-arrow {{
                /* Use standard arrows or provide custom images */
                /* image: url(path/to/up_arrow.png); */
                width: 10px; height: 10px;
            }}
            QSpinBox::down-button {{
                subcontrol-origin: border;
                subcontrol-position: bottom right; /* Position at the bottom right */
                width: 22px;
                height: 50%;
                border-left: 1px solid {theme.get('border', '#DDDDDD')};
                border-bottom-right-radius: 3px; /* Match main radius */
                background-color: {theme.get('button_bg', '#E0E0E0')};
            }}
            QSpinBox::down-button:hover {{
                background-color: {theme.get('button_hover', '#BDBDBD')};
            }}
            QSpinBox::down-arrow {{
                /* image: url(path/to/down_arrow.png); */
                 width: 10px; height: 10px;
            }}
             QSpinBox:disabled {{
                background-color: {theme.get('button_hover', '#BDBDBD')}80;
                color: {theme.get('text_secondary', '#AAAAAA')};
                border-color: {theme.get('border', '#DDDDDD')}80;
            }}
            QSpinBox::up-button:disabled, QSpinBox::down-button:disabled {{
                background-color: {theme.get('bg_secondary', '#F5F5F5')}80;
            }}
        """)


class StyledComboBox(QComboBox):
    """Custom styled combo box (dropdown)."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        # Initial style (will be overridden by setTheme)
        self.setStyleSheet("QComboBox { border: 1px solid #DDDDDD; border-radius: 4px; padding: 8px; min-width: 100px; }")

        # --- CORRECTED LINE ---
        # Use a concrete view like QListView for the popup
        self.setView(QListView(self))
        # ---------------------

    def setTheme(self, theme: ThemeDict):
        """Applies theme colors to the combo box and its dropdown."""
        self.setStyleSheet(f"""
            QComboBox {{
                border: 1px solid {theme.get('border', '#DDDDDD')};
                border-radius: 4px;
                padding: 8px;
                padding-right: 25px; /* Make space for the dropdown arrow */
                background-color: {theme.get('bg_secondary', '#F5F5F5')};
                color: {theme.get('text_primary', '#333333')};
                min-width: 100px; /* Ensure minimum width */
                min-height: 20px; /* Ensure minimum height */
            }}
            QComboBox:focus {{
                border: 1px solid {theme.get('accent', '#1976D2')};
            }}
            QComboBox:on {{ /* Styles when the dropdown is open */
                border: 1px solid {theme.get('accent', '#1976D2')};
            }}
            QComboBox::drop-down {{ /* The dropdown button */
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 22px;
                border-left-width: 1px;
                border-left-color: {theme.get('border', '#DDDDDD')};
                border-left-style: solid;
                border-top-right-radius: 3px; /* Match main radius */
                border-bottom-right-radius: 3px;
                background-color: {theme.get('button_bg', '#E0E0E0')};
            }}
            QComboBox::drop-down:hover {{
                 background-color: {theme.get('button_hover', '#BDBDBD')};
            }}
            QComboBox::down-arrow {{
                /* Standard arrow is usually fine */
                width: 10px; height: 10px;
            }}

            /* Style the popup list view */
            QComboBox QAbstractItemView {{ /* Target the view *within* the combo box */
                border: 1px solid {theme.get('border', '#DDDDDD')};
                border-radius: 4px;
                background-color: {theme.get('bg_secondary', '#F5F5F5')};
                color: {theme.get('text_primary', '#333333')};
                padding: 4px;
                outline: 0px; /* Remove focus outline from popup */
                /* Selection colors are handled by the view's palette/stylesheet */
            }}
            QComboBox QAbstractItemView::item {{
                 padding: 5px;
                 min-height: 25px;
            }}
            /* Explicitly style the selection within the list view */
            QComboBox QAbstractItemView::item:selected {{
                background-color: {theme.get('accent', '#1976D2')};
                color: white;
            }}
            QComboBox QAbstractItemView::item:hover:!selected {{
                background-color: {theme.get('button_hover', '#BDBDBD')}30; /* Subtle hover */
                color: {theme.get('text_primary', '#333333')};
            }}

            /* Disabled state */
            QComboBox:disabled {{
                background-color: {theme.get('button_hover', '#BDBDBD')}80;
                color: {theme.get('text_secondary', '#AAAAAA')};
                border-color: {theme.get('border', '#DDDDDD')}80;
            }}
             QComboBox::drop-down:disabled {{
                background-color: {theme.get('bg_secondary', '#F5F5F5')}80;
            }}
        """)


# --- Dashboard Components ---

class MetricCard(QFrame):
    """A styled frame to display a single metric (title and value)."""
    def __init__(self, title: str, initial_value: str, theme: ThemeDict, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setFrameShape(QFrame.Shape.StyledPanel) # Use styled panel shape

        self.title_label = QLabel(title)
        self.value_label = QLabel(initial_value)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(5) # Less spacing between title and value
        layout.addWidget(self.title_label)
        layout.addWidget(self.value_label)
        layout.addStretch(1) # Push content to top

        self.setTheme(theme) # Apply initial theme

    def setValue(self, value: str):
        """Updates the displayed metric value."""
        self.value_label.setText(str(value)) # Ensure value is string

    def setTheme(self, theme: ThemeDict):
        """Applies theme colors to the metric card."""
        self.setStyleSheet(f"""
            MetricCard {{ /* Use class name for specificity */
                border: 1px solid {theme.get('border', '#DDDDDD')};
                border-radius: 6px;
                background-color: {theme.get('bg_primary', '#FFFFFF')};
                padding: 15px; /* Handled by layout margins now */
                min-height: 80px; /* Ensure cards have some minimum height */
            }}
        """)
        self.title_label.setFont(QFont("Segoe UI", 9))
        self.title_label.setStyleSheet(f"color: {theme.get('text_secondary', '#666666')}; border: none; background: transparent;")
        self.value_label.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        self.value_label.setStyleSheet(f"color: {theme.get('accent', '#1976D2')}; border: none; background: transparent;")


class DashboardWidget(QWidget): # Keep as QWidget, styling is on sub-components
    """Widget containing multiple metric cards and potentially charts."""

    def __init__(self, theme: ThemeDict, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.theme = theme
        # Initialize UI elements before calling setTheme
        self.email_card = None
        self.checked_card = None
        self.latest_card = None
        self.chart_frame = None
        self.chart_label = None
        self._init_ui()

    def _init_ui(self):
        """Initializes the UI components of the dashboard."""
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0) # No margins for the main container
        main_layout.setSpacing(15)

        # --- Metrics Grid ---
        grid_layout = QHBoxLayout()
        grid_layout.setSpacing(15)

        # Create metric cards (pass the initial theme)
        self.email_card = MetricCard("Total Matched Emails", "0", self.theme)
        self.checked_card = MetricCard("Processed Today", "0", self.theme) # Changed title
        self.latest_card = MetricCard("Latest Match", "Never", self.theme) # Changed title

        grid_layout.addWidget(self.email_card)
        grid_layout.addWidget(self.checked_card)
        grid_layout.addWidget(self.latest_card)
        main_layout.addLayout(grid_layout)

        # --- Chart Placeholder ---
        self.chart_frame = QFrame()
        self.chart_frame.setFrameShape(QFrame.Shape.StyledPanel)
        self.chart_frame.setMinimumHeight(200)

        chart_layout = QVBoxLayout(self.chart_frame)
        chart_layout.setContentsMargins(15,15,15,15)
        self.chart_label = QLabel("Activity Chart (Placeholder)")
        self.chart_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.chart_label.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
        chart_layout.addWidget(self.chart_label)
        # TODO: Add actual chart widget here later if needed

        main_layout.addWidget(self.chart_frame, stretch=1) # Allow chart frame to stretch

        # Apply theme to chart frame and label initially
        self.setTheme(self.theme) # Apply theme to components created here

    def update_metrics(self, total_emails: int, processed_count: int, latest_match_time: str):
        """Updates the values displayed on the metric cards."""
        if self.email_card: self.email_card.setValue(str(total_emails))
        if self.checked_card: self.checked_card.setValue(str(processed_count))
        if self.latest_card: self.latest_card.setValue(latest_match_time if latest_match_time else "Never")

    def setTheme(self, theme: ThemeDict):
        """Updates the theme for the dashboard and its components."""
        self.theme = theme
        # Update theme for components that need it
        if self.email_card: self.email_card.setTheme(theme)
        if self.checked_card: self.checked_card.setTheme(theme)
        if self.latest_card: self.latest_card.setTheme(theme)

        # Update chart frame style
        if self.chart_frame:
            self.chart_frame.setStyleSheet(f"""
                QFrame {{
                    border: 1px solid {theme.get('border', '#DDDDDD')};
                    border-radius: 6px;
                    background-color: {theme.get('bg_secondary', '#F5F5F5')};
                }}
            """)
        # Update chart label color
        if self.chart_label:
            self.chart_label.setStyleSheet(f"color: {theme.get('text_primary', '#333333')}; border: none; background: transparent;")