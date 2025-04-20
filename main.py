# -*- coding: utf-8 -*-
import sys
import os

# --- Early Check for Essential Libraries ---
# Do this before importing other modules that might fail cryptically if PyQt6 etc. are missing
try:
    from PyQt6.QtWidgets import QApplication, QSplashScreen, QLabel, QProgressBar, QVBoxLayout
    from PyQt6.QtGui import QPixmap, QFont, QColor
    from PyQt6.QtCore import Qt
    # Check for other absolutely critical non-standard libraries if needed upfront
    # import plyer # Example - less critical, can fail later
except ImportError as e:
     print(f"Error: Missing essential library: {e.name}")
     print("-------------------------------------------------------")
     print("Please ensure you have installed all requirements.")
     print("Activate your virtual environment and run:")
     print("  pip install -r requirements.txt")
     print("-------------------------------------------------------")
     sys.exit(1) # Exit early if core libs are missing

# --- Set High DPI Attributes (BEFORE QApplication creation) ---
# This can improve rendering on high-resolution displays
try:
    if hasattr(Qt.ApplicationAttribute, 'AA_EnableHighDpiScaling'):
        QApplication.setAttribute(Qt.ApplicationAttribute.AA_EnableHighDpiScaling, True)
    if hasattr(Qt.ApplicationAttribute, 'AA_UseHighDpiPixmaps'):
        QApplication.setAttribute(Qt.ApplicationAttribute.AA_UseHighDpiPixmaps, True)
except Exception as e:
    print(f"Warning: Could not set High DPI attributes: {e}")

# --- Import the main application window ---
# Do this *after* the initial library check
try:
    from gmail_attachment_monitor.main_window import GmailMonitor
except ImportError as e:
    print(f"Error: Could not import application components: {e}")
    print("Ensure the project structure is correct and you are running")
    print("python main.py from the 'gmail-attachment-monitor' directory.")
    sys.exit(1)


# === Application Runner Class ===
class GmailMonitorApp:
    """Handles application initialization, splash screen, and execution."""

    def __init__(self):
        self.app = QApplication(sys.argv)
        self.app.setStyle('Fusion') # Good default cross-platform style

        self.splash = None
        self.window = None

        try:
            # Show splash screen
            self.splash = self.create_splash_screen()
            if self.splash:
                self.splash.show()
                # Process events to make sure splash screen is displayed immediately
                self.app.processEvents()
                self.splash.showMessage("Initializing components...",
                                       Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignCenter,
                                       QColor("white")) # Example message update

            # Create and prepare the main window
            # Initialization errors within GmailMonitor should be handled there (e.g., show QMessageBox)
            self.window = GmailMonitor()

            # Show main window (it might take a moment to fully render)
            self.window.show()
            self.app.processEvents() # Help process window show events

            # Close splash screen once the main window is up
            if self.splash:
                self.splash.finish(self.window)

            # Start the application event loop
            sys.exit(self.app.exec())

        except Exception as e:
            # Catch unexpected critical errors during startup
            print(f"FATAL ERROR during application startup: {e}")
            # Attempt to show a simple message box if possible
            try:
                from PyQt6.QtWidgets import QMessageBox
                msgBox = QMessageBox()
                msgBox.setIcon(QMessageBox.Icon.Critical)
                msgBox.setWindowTitle("Fatal Startup Error")
                msgBox.setText(f"A critical error occurred during startup:\n\n{e}\n\nThe application will now exit.")
                msgBox.setStandardButtons(QMessageBox.StandardButton.Ok)
                msgBox.exec()
            except Exception:
                # If even QMessageBox fails, just exit
                pass
            sys.exit(1)


    def create_splash_screen(self) -> QSplashScreen | None:
        """Creates and configures the splash screen."""
        try:
            pixmap_width, pixmap_height = 450, 250
            # Use a simple color fill or load an image QPixmap('path/to/splash.png')
            splash_pix = QPixmap(pixmap_width, pixmap_height)
            splash_pix.fill(QColor("#E3F2FD")) # Light Blue theme background

            splash = QSplashScreen(splash_pix, Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.FramelessWindowHint)
            splash.setStyleSheet("""
                QSplashScreen {
                    /* Optional: add border */
                    border: 1px solid #1976D2;
                    border-radius: 5px;
                    background-color: #E3F2FD; /* Ensure BG color matches pixmap */
                }
            """)

            # Use a layout to position elements on the splash screen
            splash_layout = QVBoxLayout(splash) # Apply layout directly to splash
            splash_layout.setContentsMargins(20, 20, 20, 20)
            splash_layout.addStretch(1) # Push content down

            title_label = QLabel("Gmail Attachment Monitor")
            title_label.setFont(QFont("Segoe UI", 22, QFont.Weight.Bold))
            # Ensure label background is transparent if pixmap/border is used
            title_label.setStyleSheet("color: #1A237E; background-color: transparent; border: none;")
            title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            splash_layout.addWidget(title_label)

            # No need for subtitle label here if using splash.showMessage

            splash_layout.addSpacing(15)

            progress = QProgressBar()
            progress.setRange(0, 0) # Indeterminate
            progress.setTextVisible(False)
            progress.setStyleSheet("""
                QProgressBar {
                    border: none;
                    border-radius: 4px;
                    background-color: #BBDEFB; /* Lighter blue */
                    height: 8px;
                    margin-left: 30px; /* Add margins */
                    margin-right: 30px;
                }
                QProgressBar::chunk {
                    background-color: #1976D2; /* Accent blue */
                    border-radius: 4px;
                }
            """)
            splash_layout.addWidget(progress)
            splash_layout.addStretch(2) # More space at the bottom

            return splash
        except Exception as e:
            print(f"Warning: Could not create splash screen: {e}")
            return None


# === Main Execution Block ===
if __name__ == "__main__":
    # Set environment variables for robust IO encoding, especially on Windows
    try:
        os.environ['PYTHONIOENCODING'] = 'utf-8'
        os.environ['PYTHONLEGACYWINDOWSSTDIO'] = 'utf-8'
    except Exception as e:
        print(f"Warning: Could not set environment variables for IO encoding: {e}")

    # Instantiate and run the app
    GmailMonitorApp()