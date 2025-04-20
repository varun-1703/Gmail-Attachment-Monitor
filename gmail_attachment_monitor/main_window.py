# -*- coding: utf-8 -*-
# Keep the existing imports, adjust paths for moved components
import os
import time
import base64
import datetime
import threading
import sys
import pickle
import io
import re
import PyPDF2
import docx
import pandas as pd
import zipfile
import datetime
import traceback 
from .utils import escape_html 

from bs4 import BeautifulSoup
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from plyer import notification # Keep plyer import here if send_notification stays here

from PyQt6.QtWidgets import (QMainWindow, QTabWidget, QWidget, QVBoxLayout,
                             QHBoxLayout, QLabel, QTextEdit, QTreeWidget,
                             QTreeWidgetItem, QGroupBox, QSplitter, QMessageBox,
                             QHeaderView, QFormLayout, QFrame, QStatusBar) # Removed QApplication, QSplashScreen, QProgressBar, QPushButton etc. - they are in widgets or main
from PyQt6.QtCore import Qt, pyqtSignal, QTimer # Keep QThread related if threads stay as inner classes, otherwise remove
from PyQt6.QtGui import QFont, QColor, QIcon, QPixmap # Keep QFont, QColor etc.

# Import components from other files in the package
from .theme import ThemeManager
from .widgets import (StyledProgressBar, PulseProgressBar, StyledButton, StyledGroupBox,
                      StyledTreeWidget, StyledTextEdit, StyledLineEdit, StyledSpinBox,
                      StyledComboBox, DashboardWidget) # Import necessary widgets
from .threads import EmailCheckerThread, ManualCheckThread # Import threads
from .utils import (get_email_body, format_sender, escape_html,
                    format_body_text_no_highlight, get_attachments_info) # Import utils

# --- GmailMonitor Class ---
# (Copy the ENTIRE GmailMonitor class here)
# - Remove the ThemeManager class from here (it's imported)
# - Remove the Styled* widget classes (they are imported)
# - Remove the EmailCheckerThread and ManualCheckThread classes (they are imported)
# - Remove the DashboardWidget class (it's imported)
# - Remove the helper functions that were moved to utils.py (they are imported)
# - Update references to use imported classes/functions (e.g., `self.theme_combo = StyledComboBox()`)
# - Update imports within this class if necessary (most should be covered by top-level imports)
# - Decide if download_attachment and extract_text_from_attachment remain here or move to utils.py
#   (Keeping them here is simpler for accessing 'self.service' and 'self.log')

class GmailMonitor(QMainWindow):
    def __init__(self):
        super().__init__()

        # Gmail API settings (keep as is)
        self.SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
        self.creds = None
        self.service = None

        # Application state (keep as is)
        self.running = False
        self.search_term = "varun" # Default keyword
        self.days_to_search = 1
        self.check_interval = 300  # 5 minutes
        self.monitor_thread = None
        self.manual_thread = None # Add reference for manual thread
        self.theme = ThemeManager.THEMES["Light"] # Default theme

        # Store matched emails (keep as is)
        self.matched_emails = []
        self.processed_ids = set()

        # UI Initialization is done via methods
        self.init_ui()

        # Attempt to connect to Gmail API after UI is set up
        self.connect_to_gmail() # Ensure connection happens after UI is ready for logging

    def init_ui(self):
        """Initialize the user interface"""
        self.setWindowTitle("Gmail Attachment Monitor")
        self.setGeometry(100, 100, 1200, 800) # Adjusted default size

        app_font = QFont("Segoe UI", 10)
        self.setFont(app_font)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(10)

        # --- Header ---
        header_layout = QHBoxLayout()
        logo_label = QLabel("Gmail Attachment Monitor")
        logo_label.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        header_layout.addWidget(logo_label)
        header_layout.addStretch()
        theme_label = QLabel("Theme:")
        header_layout.addWidget(theme_label)
        self.theme_combo = StyledComboBox() # Use imported widget
        self.theme_combo.addItems(list(ThemeManager.THEMES.keys()))
        self.theme_combo.setCurrentText("Light")
        self.theme_combo.currentTextChanged.connect(self.change_theme)
        header_layout.addWidget(self.theme_combo)
        main_layout.addLayout(header_layout)

        # --- Tab Widget ---
        self.tab_widget = QTabWidget()
        main_layout.addWidget(self.tab_widget)

        # Create tabs (add Dashboard tab potentially)
        # self.dashboard_tab = QWidget() # Optional: Add dashboard
        self.monitoring_tab = QWidget()
        self.summaries_tab = QWidget()

        # self.tab_widget.addTab(self.dashboard_tab, "Dashboard")
        self.tab_widget.addTab(self.monitoring_tab, "Monitoring Settings")
        self.tab_widget.addTab(self.summaries_tab, "Matched Emails (0)")

        # Setup tab content
        # self.setup_dashboard_tab() # Optional
        self.setup_monitoring_tab()
        self.setup_summaries_tab()

        # --- Status Bar ---
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_label = QLabel("Ready")
        self.status_bar.addWidget(self.status_label, 1) # Give label stretch factor
        self.progress_bar = PulseProgressBar() # Use imported widget
        self.status_bar.addPermanentWidget(self.progress_bar, 1) # Give progress bar stretch factor

        # Apply initial theme
        self.change_theme("Light")

    # --- setup_dashboard_tab (Optional) ---
    # def setup_dashboard_tab(self):
    #    layout = QVBoxLayout(self.dashboard_tab)
    #    # Pass the theme dict to DashboardWidget constructor
    #    self.dashboard_widget = DashboardWidget(self.theme)
    #    layout.addWidget(self.dashboard_widget)

    def setup_monitoring_tab(self):
        """Set up the monitoring tab UI"""
        # ... (Keep implementation, ensure Styled* widgets are used) ...
        # Example: self.search_term_input = StyledLineEdit(self.search_term)
        # Example: self.start_button = StyledButton("Start Monitoring", primary=True)
        layout = QVBoxLayout(self.monitoring_tab)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(15)

        settings_group = StyledGroupBox("Settings")
        settings_layout = QVBoxLayout()
        settings_layout.setSpacing(15)
        form_layout = QFormLayout()
        form_layout.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        form_layout.setSpacing(10) # Reduced spacing slightly

        search_label = QLabel("Username/Keyword in Attachments:")
        search_label.setFont(QFont("Segoe UI", 10))
        self.search_term_input = StyledLineEdit(self.search_term)
        form_layout.addRow(search_label, self.search_term_input)

        days_label = QLabel("Search emails from last (days):")
        days_label.setFont(QFont("Segoe UI", 10))
        self.days_spinner = StyledSpinBox()
        self.days_spinner.setValue(self.days_to_search)
        self.days_spinner.setRange(1, 30)
        form_layout.addRow(days_label, self.days_spinner)

        interval_label = QLabel("Check Interval (seconds):")
        interval_label.setFont(QFont("Segoe UI", 10))
        self.interval_spinner = StyledSpinBox()
        self.interval_spinner.setValue(self.check_interval)
        self.interval_spinner.setRange(60, 3600) # Min 1 min, Max 1 hour
        self.interval_spinner.setSingleStep(60) # Step by minute
        form_layout.addRow(interval_label, self.interval_spinner)
        settings_layout.addLayout(form_layout)

        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        self.start_button = StyledButton("Start Monitoring", primary=True)
        self.start_button.clicked.connect(self.start_monitoring)
        button_layout.addWidget(self.start_button)
        self.stop_button = StyledButton("Stop Monitoring")
        self.stop_button.clicked.connect(self.stop_monitoring)
        self.stop_button.setEnabled(False)
        button_layout.addWidget(self.stop_button)
        self.check_button = StyledButton("Check Now")
        self.check_button.clicked.connect(self.check_now)
        button_layout.addWidget(self.check_button)
        settings_layout.addLayout(button_layout)
        settings_group.setLayout(settings_layout)
        layout.addWidget(settings_group)

        status_group = StyledGroupBox("Monitoring Status")
        status_layout = QVBoxLayout()
        self.status_text = QLabel("Not monitoring")
        self.status_text.setFont(QFont("Segoe UI", 10))
        self.status_text.setWordWrap(True) # Allow wrapping
        status_layout.addWidget(self.status_text)
        self.monitor_progress = StyledProgressBar() # Progress for manual check
        status_layout.addWidget(self.monitor_progress)
        status_group.setLayout(status_layout)
        layout.addWidget(status_group)

        log_group = StyledGroupBox("Activity Log")
        log_layout = QVBoxLayout()
        self.log_text = StyledTextEdit() # Read-only by default
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group, stretch=1) # Allow log to stretch


    def setup_summaries_tab(self):
        """Set up the summaries tab UI"""
        # ... (Keep implementation, ensure Styled* widgets are used) ...
        # Example: self.email_tree = StyledTreeWidget()
        layout = QVBoxLayout(self.summaries_tab)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(15)

        # --- Top controls ---
        top_layout = QHBoxLayout()
        stats_frame = QFrame() # Use standard QFrame, style applied in change_theme
        stats_layout = QHBoxLayout(stats_frame)
        self.total_emails_label = QLabel("Total Emails: 0")
        self.total_emails_label.setFont(QFont("Segoe UI", 10))
        stats_layout.addWidget(self.total_emails_label)
        stats_layout.addWidget(QLabel("|"))
        self.last_check_label = QLabel("Last Check: Never")
        self.last_check_label.setFont(QFont("Segoe UI", 10))
        stats_layout.addWidget(self.last_check_label)
        top_layout.addWidget(stats_frame)
        top_layout.addStretch()
        controls_layout = QHBoxLayout()
        refresh_button = StyledButton("Refresh")
        refresh_button.clicked.connect(self.refresh_summaries)
        controls_layout.addWidget(refresh_button)
        clear_button = StyledButton("Clear All")
        clear_button.clicked.connect(self.clear_summaries)
        controls_layout.addWidget(clear_button)
        top_layout.addLayout(controls_layout)
        layout.addLayout(top_layout)

        # --- Filter ---
        filter_layout = QHBoxLayout()
        filter_label = QLabel("Filter:")
        filter_label.setFont(QFont("Segoe UI", 10))
        filter_layout.addWidget(filter_label)
        self.filter_input = StyledLineEdit()
        self.filter_input.setPlaceholderText("Filter by sender, subject, or matched file...") # Placeholder text
        self.filter_input.textChanged.connect(self.filter_summaries)
        filter_layout.addWidget(self.filter_input)
        layout.addLayout(filter_layout)

        # --- Splitter ---
        splitter = QSplitter(Qt.Orientation.Vertical)
        splitter.setChildrenCollapsible(False)
        layout.addWidget(splitter, 1) # Allow splitter to stretch

        # --- Email list group ---
        email_list_group = StyledGroupBox("Email List (Matches in Attachments)")
        email_list_layout = QVBoxLayout()
        self.email_tree = StyledTreeWidget()
        self.email_tree.setHeaderLabels(["Timestamp", "Sender", "Subject", "Matched Attachment(s)"])
        self.email_tree.setSelectionMode(QTreeWidget.SelectionMode.SingleSelection)
        self.email_tree.itemClicked.connect(self.show_email_details)
        header = self.email_tree.header()
        header.setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents) # Timestamp, Sender, Matched Files
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch) # Subject stretches
        header.setSortIndicatorShown(True) # Show sort indicator
        self.email_tree.sortByColumn(0, Qt.SortOrder.DescendingOrder) # Default sort
        email_list_layout.addWidget(self.email_tree)
        email_list_group.setLayout(email_list_layout)
        splitter.addWidget(email_list_group)

        # --- Email details group ---
        self.details_group = StyledGroupBox("Email Details")
        details_layout = QVBoxLayout()
        self.details_text = StyledTextEdit() # Read-only by default
        self.details_text.setReadOnly(True)
        details_layout.addWidget(self.details_text)
        self.details_group.setLayout(details_layout)
        splitter.addWidget(self.details_group)

        # Set initial splitter sizes (adjust as needed)
        splitter.setSizes([self.height() // 2, self.height() // 2]) # Equal split initially


    # --- connect_to_gmail ---
    def connect_to_gmail(self):
        """Connect to Gmail API"""
        # ... (Keep full implementation) ...
        # Ensure 'credentials.json' and 'token.pickle' paths are correct (should be root)
        creds_path = 'credentials.json'
        token_path = 'token.pickle'
        try:
            if os.path.exists(token_path):
                with open(token_path, 'rb') as token:
                    self.creds = pickle.load(token)

            if not self.creds or not self.creds.valid:
                if self.creds and self.creds.expired and self.creds.refresh_token:
                    try:
                        self.creds.refresh(Request())
                    except Exception as refresh_err:
                         self.log(f"Error refreshing token: {refresh_err}. Please re-authenticate.", "error")
                         # Force re-authentication by deleting token file
                         if os.path.exists(token_path):
                             os.remove(token_path)
                         self.creds = None # Reset creds
                         # Re-run the flow below
                # else: # Covers no creds, or creds without refresh token
                if not self.creds: # Re-check after potential refresh failure/deletion
                    if not os.path.exists(creds_path):
                        self.log(f"Error: {creds_path} not found. Please download it from Google Cloud Console and place it in the application's root directory.", "error")
                        QMessageBox.critical(self, "Credentials Error", f"{creds_path} not found. See README for setup instructions.")
                        # Disable buttons or exit? For now, just log and show message.
                        self.start_button.setEnabled(False)
                        self.check_button.setEnabled(False)
                        return # Stop connection attempt

                    flow = InstalledAppFlow.from_client_secrets_file(creds_path, self.SCOPES)
                    # Make run_local_server slightly more robust
                    try:
                         self.creds = flow.run_local_server(port=0, prompt='consent', authorization_prompt_message='Please authorize access to Gmail Read-Only:\n')
                    except Exception as flow_err:
                         self.log(f"Error during OAuth flow: {flow_err}", "error")
                         QMessageBox.critical(self, "Authentication Error", f"Failed to complete authentication: {flow_err}")
                         return

                # Save the credentials for the next run
                try:
                    with open(token_path, 'wb') as token:
                        pickle.dump(self.creds, token)
                except Exception as dump_err:
                    self.log(f"Warning: Could not save token to {token_path}: {dump_err}", "warning")


            # Build the service object if creds are valid
            if self.creds and self.creds.valid:
                 self.service = build('gmail', 'v1', credentials=self.creds)
                 self.log("Successfully connected to Gmail API", "success")
                 # Enable buttons now that connection is successful
                 if not self.running: # Only enable if not already monitoring
                     self.start_button.setEnabled(True)
                     self.check_button.setEnabled(True)
            else:
                 self.log("Failed to obtain valid credentials after OAuth flow.", "error")
                 QMessageBox.critical(self, "Authentication Error", "Could not obtain valid credentials. Please check console logs and ensure you completed the authorization.")
                 self.start_button.setEnabled(False)
                 self.check_button.setEnabled(False)


        except Exception as e:
            self.log(f"Failed to connect to Gmail API: {str(e)}", "error")
            QMessageBox.critical(self, "Connection Error",
                                f"Failed to connect to Gmail API: {str(e)}\nPlease check your internet connection and {creds_path}.")
            self.start_button.setEnabled(False)
            self.check_button.setEnabled(False)


    # --- start_monitoring ---
    def start_monitoring(self):
        """Start monitoring emails"""
        # ... (Keep implementation) ...
        # Use imported EmailCheckerThread
        if not self.service:
             QMessageBox.warning(self, "Error", "Cannot start monitoring. Not connected to Gmail API. Check logs.")
             return

        self.search_term = self.search_term_input.text().strip()
        self.days_to_search = self.days_spinner.value()
        self.check_interval = self.interval_spinner.value()

        if not self.search_term:
            QMessageBox.warning(self, "Input Error", "Please enter a keyword to search for in attachments.")
            return

        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.check_button.setEnabled(False)
        self.search_term_input.setEnabled(False)
        self.days_spinner.setEnabled(False)
        self.interval_spinner.setEnabled(False)

        status_msg = f"Monitoring attachments for: '{self.search_term}' every {self.check_interval}s"
        self.status_text.setText(status_msg)
        self.status_label.setText(f"Monitoring active | Keyword: '{self.search_term}'")
        self.progress_bar.start_pulse()
        self.monitor_progress.setValue(0) # Reset manual progress bar

        self.running = True
        self.monitor_thread = EmailCheckerThread(self) # Pass self
        self.monitor_thread.log_signal.connect(self.log)
        self.monitor_thread.update_signal.connect(self.update_ui) # If thread emits update signal
        # self.monitor_thread.progress_signal.connect(self.progress_bar.setValue) # Connect interval progress
        self.monitor_thread.start()

        self.log(status_msg, "info")


    # --- stop_monitoring ---
    def stop_monitoring(self):
        """Stop monitoring emails"""
        # ... (Keep implementation) ...
        self.running = False
        if self.monitor_thread and self.monitor_thread.isRunning():
            self.log("Stopping monitoring thread...", "info")
            self.monitor_thread.running = False # Signal thread to stop
            self.monitor_thread.quit() # Request event loop exit
            if not self.monitor_thread.wait(5000): # Wait up to 5s
                 self.log("Monitoring thread did not stop gracefully, terminating.", "warning")
                 self.monitor_thread.terminate()
                 self.monitor_thread.wait() # Wait after terminate
            self.log("Monitoring stopped.", "info")

        self.monitor_thread = None

        # Update UI state
        self.start_button.setEnabled(True if self.service else False) # Only enable if connected
        self.stop_button.setEnabled(False)
        self.check_button.setEnabled(True if self.service else False) # Only enable if connected
        self.search_term_input.setEnabled(True)
        self.days_spinner.setEnabled(True)
        self.interval_spinner.setEnabled(True)

        self.status_text.setText("Monitoring stopped.")
        self.status_label.setText("Ready" if self.service else "Not Connected")
        self.monitor_progress.setValue(0)
        self.progress_bar.stop_pulse()
        self.progress_bar.setValue(0)

        # Re-log stopped message in case thread termination took time
        # self.log("Stopped monitoring", "info")


    # --- check_now ---
    def check_now(self):
        """Perform a manual check"""
        # ... (Keep implementation) ...
        # Use imported ManualCheckThread
        if not self.service:
             QMessageBox.warning(self, "Error", "Cannot check now. Not connected to Gmail API. Check logs.")
             return
        if self.manual_thread and self.manual_thread.isRunning():
             self.log("Manual check already in progress.", "warning")
             return

        self.search_term = self.search_term_input.text().strip()
        self.days_to_search = self.days_spinner.value()

        if not self.search_term:
            QMessageBox.warning(self, "Input Error", "Please enter a keyword to search for in attachments.")
            return

        self.check_button.setEnabled(False)
        self.start_button.setEnabled(False) # Disable start during manual check
        self.stop_button.setEnabled(False) # Disable stop during manual check

        self.monitor_progress.setValue(0)
        self.status_text.setText(f"Manual check running for: '{self.search_term}'...")
        self.status_label.setText("Manual check in progress...")

        # Create and start thread
        self.manual_thread = ManualCheckThread(self) # Pass self
        self.manual_thread.progress_signal.connect(self.monitor_progress.setValue)
        self.manual_thread.finished_signal.connect(self.manual_check_finished)
        self.manual_thread.start()

        self.log(f"Manual check started for attachments containing '{self.search_term}'", "info")


    # --- manual_check_finished ---
    def manual_check_finished(self):
        """Called when manual check is finished"""
        # ... (Keep implementation) ...
        self.manual_thread = None # Clear thread reference

        # Re-enable buttons appropriately
        self.check_button.setEnabled(True if self.service else False)
        self.start_button.setEnabled(True if self.service and not self.running else False)
        # Stop button remains disabled as monitoring isn't running

        self.status_text.setText("Manual check completed.")
        self.status_label.setText("Ready" if self.service else "Not Connected")
        # Optionally leave progress bar at 100 for a moment, then reset?
        # QTimer.singleShot(2000, lambda: self.monitor_progress.setValue(0))
        self.log("Manual check completed", "info")


    # --- check_emails ---
    # Keep implementation, but potentially use helper functions from utils.py
    # Ensure self.log is used for logging
    def check_emails(self, manual_thread=None):
        """Check emails for the search term within attachments."""
        # ... (Keep full implementation as in your original script) ...
        # Make sure it uses self.log, self.service, self.search_term, etc.
        # Make sure it uses the imported `get_email_body` and `get_attachments_info`
        # Make sure it calls self.download_attachment and self.extract_text_from_attachment
        if not self.service:
            self.log("Gmail service not available for check_emails.", "error")
            if manual_thread: manual_thread.progress_signal.emit(0)
            return
        try:
            now = datetime.datetime.now()
            after_date = now - datetime.timedelta(days=self.days_to_search)
            after_str = after_date.strftime('%Y/%m/%d')
            query = f'has:attachment after:{after_str}'

            self.log(f"Executing query: {query}", "info")
            request = self.service.users().messages().list(userId='me', q=query, maxResults=100) # Increase maxResults?
            response = request.execute()
            messages = response.get('messages', [])
            total_messages_to_process = len(messages)
            processed_count = 0
            page_token = response.get('nextPageToken')

            # --- Pagination Handling ---
            while page_token:
                if manual_thread is None and self.monitor_thread and not self.monitor_thread.running:
                    self.log("Monitoring stopped during pagination.", "info")
                    break # Exit pagination loop
                self.log(f"Fetching next page of results...", "info")
                request = self.service.users().messages().list(userId='me', q=query, maxResults=100, pageToken=page_token)
                response = request.execute()
                messages.extend(response.get('messages', []))
                page_token = response.get('nextPageToken')
                total_messages_to_process = len(messages) # Update total count
            # --- End Pagination ---

            if not messages:
                self.log(f"No emails with attachments found matching criteria.", "info")
                if manual_thread: manual_thread.progress_signal.emit(100)
                self.last_check_label.setText(f"Last Check: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                return

            self.log(f"Found ~{total_messages_to_process} potential emails with attachments to check.", "info")

            new_emails_matched = 0
            search_term_lower = self.search_term.lower()

            for i, message_summary in enumerate(messages):
                 # Check cancellation flag (important for responsiveness)
                 if manual_thread is None and self.monitor_thread and not self.monitor_thread.running:
                      self.log("Monitoring stopped during check.", "info")
                      break # Exit the loop
                 if manual_thread and not self.manual_thread: # Check if manual thread still exists
                      self.log("Manual check cancelled.", "info")
                      break

                 msg_id = message_summary['id']

                 # Skip if already processed OR already in matched list
                 if msg_id in self.processed_ids or any(email['id'] == msg_id for email in self.matched_emails):
                      processed_count += 1
                      if manual_thread:
                           progress = int(((processed_count) / total_messages_to_process) * 90) + 10
                           manual_thread.progress_signal.emit(progress)
                      continue

                 try:
                      # Get full message details
                      # Consider adding fields='id,payload(headers,parts,body),internalDate' to optimize?
                      msg = self.service.users().messages().get(userId='me', id=msg_id, format='full').execute()
                      payload = msg.get('payload')
                      if not payload:
                           self.log(f"No payload found for message {msg_id}", "warning")
                           self.processed_ids.add(msg_id) # Mark processed
                           processed_count += 1
                           continue # Skip to next message

                      headers = payload.get('headers', [])
                      subject = next((h['value'] for h in headers if h['name'].lower() == 'subject'), 'No Subject')
                      sender = next((h['value'] for h in headers if h['name'].lower() == 'date'), '') # Use imported util
                      date_str = next((h['value'] for h in headers if h['name'].lower() == 'date'), '')

                      # Find attachments using imported util
                      attachments = get_attachments_info(payload.get('parts'))
                      if not attachments:
                           self.log(f"Message {msg_id} had 'has:attachment' but no attachments found via parsing.", "warning")
                           self.processed_ids.add(msg_id)
                           processed_count += 1
                           continue

                      match_found_in_attachment = False
                      matched_filenames = []
                      for attachment_info in attachments:
                           filename = attachment_info['filename']
                           att_id = attachment_info['attachmentId']
                           mime_type = attachment_info['mimeType']

                           self.log(f"Checking attachment: {filename} ({mime_type}) in email {msg_id}", "debug") # Use debug level?

                           attachment_data = self.download_attachment(msg_id, att_id)
                           if not attachment_data:
                                continue

                           attachment_text = self.extract_text_from_attachment(attachment_data, mime_type, filename)

                           if search_term_lower in attachment_text.lower():
                               self.log(f"Keyword '{self.search_term}' FOUND in attachment: {filename}", "success")
                               match_found_in_attachment = True
                               matched_filenames.append(filename)
                               # break # Optional: stop checking other attachments in this email

                      if match_found_in_attachment:
                           body = get_email_body(msg) # Use imported util
                           timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

                           email_data = {
                               'id': msg_id,
                               'timestamp': timestamp,
                               'sender': sender,
                               'subject': subject,
                               'body': body,
                               'match_type': "Attachment",
                               'date': date_str,
                               'attachments_info': attachments,
                               'matched_filenames': matched_filenames
                           }

                           self.matched_emails.append(email_data)
                           new_emails_matched += 1

                      self.processed_ids.add(msg_id) # Mark as processed

                 except Exception as fetch_err:
                      self.log(f"Error processing message {msg_id}: {str(fetch_err)}", "error")
                      self.processed_ids.add(msg_id) # Mark processed even on error

                 finally:
                      processed_count += 1
                      if manual_thread:
                           progress = int(((processed_count) / total_messages_to_process) * 90) + 10
                           manual_thread.progress_signal.emit(progress)

            # --- After processing all messages ---
            if new_emails_matched > 0:
                self.log(f"Found {new_emails_matched} new emails with matching attachments.", "success")
                self.update_ui() # Update the list and counts
                self.send_notification(f"{new_emails_matched} new emails found",
                                      f"Found {new_emails_matched} emails with attachments containing '{self.search_term}'")
            elif total_messages_to_process > 0:
                 self.log(f"Checked {processed_count}/{total_messages_to_process} emails, no new matches found for '{self.search_term}'.", "info")
            else:
                 # This case was handled earlier if messages list was initially empty
                 pass

            self.last_check_label.setText(f"Last Check: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

        except Exception as e:
            self.log(f"Error checking emails: {str(e)}", "error")
            if manual_thread: manual_thread.progress_signal.emit(0)


    # --- download_attachment (Keep here or move to utils) ---
    def download_attachment(self, msg_id, attachment_id):
        """Download attachment data using attachmentId."""
        # ... (Keep full implementation) ...
        try:
            attachment = self.service.users().messages().attachments().get(
                userId='me', messageId=msg_id, id=attachment_id).execute()
            data = attachment.get('data')
            if data:
                # Add padding if needed for urlsafe_b64decode
                missing_padding = len(data) % 4
                if missing_padding:
                    data += '='* (4 - missing_padding)
                return base64.urlsafe_b64decode(data.encode('UTF-8'))
            else:
                self.log(f"No data found for attachment {attachment_id} in message {msg_id}", "warning")
                return None
        except Exception as e:
            self.log(f"Error downloading attachment {attachment_id} for message {msg_id}: {str(e)}", "error")
            return None


    # --- extract_text_from_attachment (Keep here or move to utils) ---
    def extract_text_from_attachment(self, attachment_data, mime_type, filename):
        """Extract text from various attachment types."""
        # ... (Keep full implementation, using self.log) ...
        text = ""
        try:
            # --- Plain Text ---
            if mime_type.startswith('text/plain') or filename.lower().endswith('.txt'):
                # Try common encodings
                for encoding in ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']:
                    try:
                        text = attachment_data.decode(encoding)
                        break
                    except UnicodeDecodeError:
                        continue
                else:
                     self.log(f"Could not decode text file: {filename} with common encodings.", "warning")
                     text = "[Could not decode text]"

            # --- PDF ---
            elif mime_type == 'application/pdf' or filename.lower().endswith('.pdf'):
                try:
                    with io.BytesIO(attachment_data) as f:
                        reader = PyPDF2.PdfReader(f, strict=False)
                        if reader.is_encrypted:
                             self.log(f"Skipping encrypted PDF: {filename}", "warning")
                             text = "[Encrypted PDF - Cannot Extract Text]"
                        else:
                             for page in reader.pages:
                                  try:
                                       page_text = page.extract_text()
                                       if page_text:
                                            text += page_text + "\n"
                                  except Exception as page_err:
                                       self.log(f"Error extracting text from page in PDF '{filename}': {page_err}", "warning")
                                       continue # Try next page
                             if not text and len(reader.pages) > 0:
                                  self.log(f"Could not extract text from PDF: {filename} (possibly image-based)", "warning")
                                  text = "[Could not extract text from PDF - Image Based?]"
                except PyPDF2.errors.PdfReadError as pdf_err:
                     self.log(f"Error reading PDF file '{filename}': {pdf_err}", "warning")
                     text = "[Invalid or Corrupted PDF]"
                except Exception as pdf_gen_err:
                    self.log(f"General error processing PDF '{filename}': {pdf_gen_err}", "error")
                    text = "[Error Processing PDF]"


            # --- DOCX ---
            elif mime_type == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' or filename.lower().endswith('.docx'):
                 try:
                     with io.BytesIO(attachment_data) as f:
                         doc = docx.Document(f)
                         full_text = [para.text for para in doc.paragraphs]
                         text = "\n".join(full_text)
                 except Exception as docx_err:
                      self.log(f"Error reading DOCX file '{filename}': {docx_err}", "warning")
                      text = "[Error reading DOCX file]"

            # --- CSV ---
            elif mime_type == 'text/csv' or filename.lower().endswith('.csv'):
                decoded_text = ""
                try:
                    # Try decoding first
                    for encoding in ['utf-8', 'latin-1', 'cp1252']:
                         try:
                              decoded_text = attachment_data.decode(encoding)
                              break
                         except UnicodeDecodeError:
                              continue
                    else:
                         raise ValueError("Could not decode CSV with common encodings")

                    with io.StringIO(decoded_text) as f:
                         # Read robustly, convert all columns to string before joining
                         df = pd.read_csv(f, on_bad_lines='warn', dtype=str, engine='python') # Use python engine for flexibility
                         df.fillna('', inplace=True) # Replace NaN with empty string
                         text = df.to_string(index=False, na_rep='') # Convert entire DataFrame to string
                except Exception as csv_err:
                     self.log(f"Error reading CSV file '{filename}': {csv_err}", "warning")
                     text = "[Error reading CSV file]"
                     if decoded_text: # If decode worked but pandas failed, provide raw text
                          text += "\n[Raw Decoded Content:\n" + decoded_text[:1000] + "...]" # Limit raw preview


            # --- XLSX ---
            elif mime_type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or filename.lower().endswith('.xlsx'):
                 try:
                     with io.BytesIO(attachment_data) as f:
                          xls = pd.ExcelFile(f, engine='openpyxl')
                          all_sheets_text = []
                          for sheet_name in xls.sheet_names:
                               try:
                                   # Read robustly, convert all to string
                                   df = xls.parse(sheet_name, dtype=str)
                                   df.fillna('', inplace=True)
                                   sheet_content = df.to_string(index=False, na_rep='')
                                   all_sheets_text.append(f"--- Sheet: {sheet_name} ---\n{sheet_content}")
                               except Exception as sheet_err:
                                   self.log(f"Error reading sheet '{sheet_name}' in XLSX '{filename}': {sheet_err}", "warning")
                                   all_sheets_text.append(f"--- Sheet: {sheet_name} ---\n[Error reading sheet]")

                          text = "\n\n".join(all_sheets_text)
                 except Exception as xlsx_err:
                     self.log(f"Error reading XLSX file '{filename}': {xlsx_err}", "warning")
                     text = "[Error reading XLSX file]"

            # --- ZIP ---
            elif mime_type in ['application/zip', 'application/x-zip-compressed'] or filename.lower().endswith('.zip'):
                 try:
                     with io.BytesIO(attachment_data) as f:
                         with zipfile.ZipFile(f, 'r') as zf:
                             text_parts = [f"--- ZIP Contents ({filename}) ---"]
                             for member_info in zf.infolist():
                                 # Basic info: filename and size
                                 text_parts.append(f"- {member_info.filename} ({member_info.file_size} bytes)")
                                 # Future: Add extraction and recursive search here if needed
                             text = "\n".join(text_parts)
                 except zipfile.BadZipFile:
                      self.log(f"Could not read ZIP file (corrupted?): {filename}", "warning")
                      text = "[Invalid ZIP file]"
                 except Exception as zip_err:
                      self.log(f"Error reading ZIP file '{filename}': {zip_err}", "warning")
                      text = "[Error reading ZIP file]"

            # --- Other/Unsupported ---
            else:
                # Only log if it's not a common non-text type we expect to ignore
                common_ignored = ['image/', 'application/octet-stream'] # Add more if needed
                if not any(mime_type.startswith(prefix) for prefix in common_ignored):
                     self.log(f"Skipping unsupported attachment type: {filename} ({mime_type})", "info")
                text = f"[Unsupported attachment type: {mime_type}]"

        except Exception as e:
            self.log(f"Critical error extracting text from {filename} ({mime_type}): {str(e)}", "error", exc_info=True) # Add stack trace
            text = "[Error during text extraction]"

        return text


    # --- update_ui ---
    def update_ui(self):
        """Update UI with matched emails"""
        # ... (Keep implementation) ...
        count = len(self.matched_emails)
        self.tab_widget.setTabText(2, f"Matched Emails ({count})") # Index 2 if dashboard is added

        self.refresh_summaries() # Refresh the list

        self.total_emails_label.setText(f"Total Emails: {count}")

        # Update dashboard if it exists
        # if hasattr(self, 'dashboard_widget'):
        #    checked_count = len(self.processed_ids) # Example metric
        #    last_check_time = self.last_check_label.text().replace("Last Check: ", "")
        #    self.dashboard_widget.update_metrics(count, checked_count, last_check_time)


    # --- refresh_summaries ---
    def refresh_summaries(self):
        """Refresh the summaries tab"""
        # ... (Keep implementation, use format_sender from utils) ...
        current_selection_id = None
        selected_items = self.email_tree.selectedItems()
        if selected_items:
            current_selection_id = selected_items[0].data(0, Qt.ItemDataRole.UserRole)

        self.email_tree.clear()
        self.email_tree.setSortingEnabled(False) # Disable sorting during population

        for email in self.matched_emails:
            item = QTreeWidgetItem()
            item.setText(0, email['timestamp'])
            item.setText(1, format_sender(email['sender'])) # Use imported util
            item.setText(2, email['subject'])
            filenames_str = ", ".join(email.get('matched_filenames', ['N/A']))
            item.setText(3, filenames_str)

            item.setData(0, Qt.ItemDataRole.UserRole, email['id']) # Store email ID

            # Tooltip for filenames column if it's long
            item.setToolTip(3, filenames_str)

            # Background color can be set here or in theme update
            # item.setBackground(0, QColor(self.theme["success"] + "40")) # Light success tint

            self.email_tree.addTopLevelItem(item)

        self.email_tree.setSortingEnabled(True) # Re-enable sorting

        # Restore selection if it still exists
        if current_selection_id:
            items = self.email_tree.findItems(current_selection_id, Qt.MatchFlag.MatchExactly | Qt.MatchFlag.MatchRecursive, 0) # Search by ID in data role
            if items:
                 self.email_tree.setCurrentItem(items[0])

        # Apply current filter again after refresh
        self.filter_summaries()


    # --- filter_summaries ---
    def filter_summaries(self):
        """Filter the email list based on text input"""
        # ... (Keep implementation) ...
        filter_text = self.filter_input.text().lower().strip()
        for i in range(self.email_tree.topLevelItemCount()):
            item = self.email_tree.topLevelItem(i)
            if not filter_text:
                item.setHidden(False)
            else:
                # Check timestamp(0), sender(1), subject(2), matched files(3)
                match = (filter_text in item.text(0).lower() or
                         filter_text in item.text(1).lower() or
                         filter_text in item.text(2).lower() or
                         filter_text in item.text(3).lower())
                item.setHidden(not match)


    # --- show_email_details ---
    def show_email_details(self, item):
        """Show details for selected email"""
        # ... (Keep implementation, use escape_html and format_body_text_no_highlight from utils) ...
        email_id = item.data(0, Qt.ItemDataRole.UserRole)
        email_data = next((e for e in self.matched_emails if e['id'] == email_id), None)

        if email_data:
            # Use imported utils for escaping
            subject_safe = escape_html(email_data['subject'])
            sender_safe = escape_html(email_data['sender'])
            date_safe = escape_html(email_data.get('date', 'N/A'))
            match_type_safe = escape_html(email_data.get('match_type', 'N/A'))

            details = f"""
            <style>
                h2 {{ margin-bottom: 5px; color: {self.theme['accent']}; }}
                p {{ margin-top: 2px; margin-bottom: 8px; }}
                b {{ color: {self.theme['text_primary']}; }}
                ul {{ margin-top: 0px; padding-left: 20px; }}
                li {{ margin-bottom: 3px; }}
                hr {{ border: none; border-top: 1px solid {self.theme['border']}; margin-top: 10px; margin-bottom: 10px; }}
                body {{ color: {self.theme['text_secondary']}; }} /* Default text color */
            </style>
            <h2>{subject_safe}</h2>
            <p><b>From:</b> {sender_safe}</p>
            <p><b>Date:</b> {date_safe}</p>
            <p><b>Match Type:</b> {match_type_safe}</p>
            """

            attachments = email_data.get('attachments_info', [])
            matched_files = email_data.get('matched_filenames', [])
            if attachments:
                details += "<p><b>Attachments:</b></p><ul>"
                for att in attachments:
                    filename_safe = escape_html(att['filename'])
                    size_mb = att.get('size', 0) / (1024 * 1024)
                    size_str = f"{size_mb:.2f} MB" if size_mb >= 0.1 else f"{att.get('size', 0) / 1024:.1f} KB"

                    if att['filename'] in matched_files:
                        # Highlight matching files clearly
                        details += f'<li><b style="color: {self.theme["success"]};">{filename_safe} (Match Found)</b> ({size_str})</li>'
                    else:
                        details += f'<li>{filename_safe} ({size_str})</li>'
                details += "</ul>"

            details += "<hr>"
            # Use imported util for body formatting
            formatted_body = format_body_text_no_highlight(email_data.get('body', ''))
            details += f"<p><b>Body Preview:</b><br>{formatted_body}</p>" if formatted_body else "<p><i>No body preview available.</i></p>"

            self.details_text.setHtml(details)
        else:
            self.details_text.setPlainText("Email details not found.")

    # --- clear_summaries ---
    def clear_summaries(self):
        """Clear all matched emails"""
        # ... (Keep implementation) ...
        reply = QMessageBox.question(
            self, "Confirm Clear",
            "Are you sure you want to clear all matched emails from the list?\n(This does not delete emails from Gmail).",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No # Default to No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self.matched_emails = []
            self.processed_ids = set() # Also clear processed IDs to allow re-checking
            self.email_tree.clear()
            self.details_text.clear()
            self.total_emails_label.setText("Total Emails: 0")
            self.tab_widget.setTabText(2, "Matched Emails (0)") # Adjust index if needed
            self.log("Cleared all matched emails and processed IDs list.", "info")


    # --- send_notification ---
    def send_notification(self, title, message):
        """Send a system notification"""
        # ... (Keep implementation) ...
        try:
            # Use plyer for cross-platform notifications
            notification.notify(
                title=title,
                message=message,
                app_name="Gmail Monitor", # Keep consistent
                timeout=15 # Longer timeout
            )
            self.log(f"Sent notification: '{title}'", "info")
        except Exception as e:
            # Log specific error if plyer backend fails
            self.log(f"Failed to send notification: {str(e)}", "warning")
            # Fallback: Show a message box? Only if critical.
            # QMessageBox.information(self, "Notification Failed", f"Could not show system notification:\n{title}\n\n{message}")


    # --- log ---
    # Keep implementation, ensure escape_html is used
    def log(self, message, level="info", exc_info=False):
        """Add message to log with appropriate color coding and timestamp."""
        try: # Add a try/except block for robustness within logging itself
            timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            level_upper = level.upper()

            # Determine color based on theme and level
            color = self.theme.get('text_primary', '#333333') # Default
            if level == "error":
                color = self.theme.get('error', '#F44336')
            elif level == "warning":
                color = self.theme.get('warning', '#FF9800')
            elif level == "success":
                color = self.theme.get('success', '#4CAF50')
            elif level == "debug": # Add debug level
                 color = self.theme.get('text_secondary', '#666666')
                 level_upper = "DEBUG" # Display as DEBUG

            # Escape the user-provided message content for safety
            safe_message = escape_html(str(message)) # Ensure message is string

            # Format the log entry as an HTML string
            formatted_entry = f"<span style='color:{color};'>[{timestamp}] {level_upper}: {safe_message}</span>"

            # Add traceback if requested (usually for errors)
            if exc_info:
                 # Use html.escape on the traceback as well
                 tb_lines = traceback.format_exc().splitlines()
                 safe_tb = escape_html("\n".join(tb_lines))
                 # Use <pre> for preformatted text (maintains spacing)
                 formatted_entry += f"<br/><pre style='color:{color}; font-size: 9px; margin-left: 10px; white-space: pre-wrap; word-wrap: break-word;'>{safe_tb}</pre>"


            # --- CORRECTED LINE ---
            # Use append() which handles HTML automatically for QTextEdit
            self.log_text.append(formatted_entry)
            # --------------------

            # Auto-scroll to the bottom (ensure log_text widget exists)
            if hasattr(self, 'log_text') and self.log_text:
                self.log_text.verticalScrollBar().setValue(self.log_text.verticalScrollBar().maximum())

        except Exception as log_err:
            # Fallback if logging itself fails
            print(f"!!! Logging Error: Could not log message '{message}'. Error: {log_err} !!!")
            # Also print original traceback if available
            if exc_info:
                 print("--- Original Traceback for Logging Failure ---")
                 traceback.print_exc()
                 print("---------------------------------------------")


    # --- change_theme ---
    def change_theme(self, theme_name):
        """Change the application theme"""
        # ... (Keep implementation, use ThemeManager.apply_theme) ...
        # Ensure self.theme is updated
        self.theme = ThemeManager.apply_theme(self.window(), theme_name) # Apply to main window's app instance

        # Update styles of all custom components
        self.update_component_styles()

        # Update specific non-custom-widget styles
        tab_style = f"""
            QTabWidget::pane {{
                border: 1px solid {self.theme["border"]};
                border-radius: 4px;
                padding: 5px;
                background-color: {self.theme["bg_primary"]}; /* Ensure pane bg matches */
            }}
            QTabBar::tab {{
                background-color: {self.theme["button_bg"]};
                color: {self.theme["text_primary"]};
                padding: 10px 20px;
                margin-right: 2px; /* Spacing between tabs */
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                border: 1px solid {self.theme["border"]}; /* Add border for definition */
                border-bottom: none; /* Remove bottom border for non-selected */
            }}
            QTabBar::tab:selected {{
                background-color: {self.theme["bg_primary"]}; /* Match pane */
                color: {self.theme["accent"]}; /* Highlight selected tab text */
                border-bottom: 1px solid {self.theme["bg_primary"]}; /* Blend bottom border */
                font-weight: bold;
            }}
            QTabBar::tab:hover:!selected {{
                background-color: {self.theme["button_hover"]};
            }}
            QTabBar::tab:disabled {{
                 background-color: {self.theme.get('bg_secondary', '#F5F5F5')};
                 color: {self.theme.get('text_secondary', '#AAAAAA')};
            }}
        """
        self.tab_widget.setStyleSheet(tab_style)

        stats_frame = self.summaries_tab.findChild(QFrame)
        if stats_frame:
             stats_frame_style = f"""
                 QFrame {{
                     border: 1px solid {self.theme["border"]};
                     border-radius: 4px;
                     background-color: {self.theme["bg_secondary"]};
                     padding: 8px 12px; /* Adjusted padding */
                 }}
                 QLabel {{ /* Style labels within this frame */
                      color: {self.theme['text_secondary']};
                      background-color: transparent; /* Ensure no bg override */
                      border: none; /* Ensure no border override */
                 }}
             """
             stats_frame.setStyleSheet(stats_frame_style)

        # Update dashboard theme if it exists
        # if hasattr(self, 'dashboard_widget'):
        #     self.dashboard_widget.set_theme(self.theme)

        # Refresh log colors if necessary (might require re-logging or storing raw logs)
        # Simplest is new logs will have the new theme colors.


    # --- update_component_styles ---
    def update_component_styles(self):
        """Update styles for all custom styled components"""
        # ... (Keep implementation) ...
        # Iterate through children and call setTheme
        for widget_type in (StyledButton, StyledGroupBox, StyledTreeWidget,
                            StyledTextEdit, StyledLineEdit, StyledSpinBox,
                            StyledComboBox, PulseProgressBar, StyledProgressBar):
            for widget in self.findChildren(widget_type):
                if hasattr(widget, 'setTheme'):
                    widget.setTheme(self.theme)

        # Manually update specific StyledProgressBar if needed (like monitor_progress)
        if hasattr(self, 'monitor_progress'):
             self.monitor_progress.setStyleSheet(f"""
                 QProgressBar {{
                     border: none;
                     border-radius: 4px;
                     background-color: {self.theme["bg_secondary"]};
                     height: 8px;
                     text-align: center; /* Center text if visible */
                 }}
                 QProgressBar::chunk {{
                     background-color: {self.theme["accent"]};
                     border-radius: 4px;
                 }}
             """)

    # --- format_sender (Moved to utils) ---
    # --- escape_html (Moved to utils) ---
    # --- format_body_text_no_highlight (Moved to utils) ---
    # --- get_email_body (Moved to utils) ---
    # --- get_attachments_info (Moved to utils) ---


    # Add closeEvent handler
    def closeEvent(self, event):
        """Handle window close event."""
        self.log("Close event triggered.", "info")
        self.stop_monitoring() # Ensure monitoring stops cleanly

        # Persist state if needed (e.g., last search term, matched emails?)
        # For simplicity, we don't save state here, but you could add pickle saves.

        event.accept() # Accept the close event