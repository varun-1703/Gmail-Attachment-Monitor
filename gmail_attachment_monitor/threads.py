# -*- coding: utf-8 -*-
import time
from PyQt6.QtCore import QThread, pyqtSignal

# --- Threads ---
# (Copy the EmailCheckerThread and ManualCheckThread classes here)

class EmailCheckerThread(QThread):
    log_signal = pyqtSignal(str, str)
    update_signal = pyqtSignal()
    progress_signal = pyqtSignal(int) # For interval progress visualization

    # ... (keep full implementation) ...

    def __init__(self, parent_monitor=None): # Changed arg name for clarity
        super().__init__(parent_monitor) # Pass parent if needed, e.g. for accessing settings directly
        self.parent_monitor = parent_monitor
        self.running = False

    def run(self):
        self.running = True
        while self.running:
            try:
                # Call the parent's method to check emails
                self.parent_monitor.check_emails()

                # Sleep for the specified interval
                total_seconds = self.parent_monitor.check_interval
                for i in range(total_seconds):
                    if not self.running:
                        break
                    # Update progress (optional visualization of interval passing)
                    progress = int(((i + 1) / total_seconds) * 100)
                    self.progress_signal.emit(progress)
                    time.sleep(1)

            except Exception as e:
                # Log error using parent's log method via signal
                self.log_signal.emit(f"Error in monitoring loop: {str(e)}", "error")
                # Avoid busy-waiting on continuous errors
                if self.running: # Check if still supposed to be running
                    time.sleep(60) # Wait before retrying

            finally:
                 # Ensure progress resets if loop finishes or errors out
                 if self.running:
                     self.progress_signal.emit(0)


class ManualCheckThread(QThread):
    progress_signal = pyqtSignal(int)
    finished_signal = pyqtSignal()

    # ... (keep full implementation) ...

    def __init__(self, parent_monitor=None): # Changed arg name for clarity
        super().__init__(parent_monitor)
        self.parent_monitor = parent_monitor

    def run(self):
        try:
            # Fake progress for UX
            for i in range(10):
                self.progress_signal.emit(i)
                time.sleep(0.01)

            # Actual check - call parent's method
            self.parent_monitor.check_emails(manual_thread=self) # Pass self

            self.progress_signal.emit(100) # Signal completion

        except Exception as e:
            # Use parent's logging via signal if preferred, or direct call if simple
            self.parent_monitor.log(f"Error during manual check: {str(e)}", "error")
            self.progress_signal.emit(0) # Reset progress on error

        finally:
            self.finished_signal.emit()