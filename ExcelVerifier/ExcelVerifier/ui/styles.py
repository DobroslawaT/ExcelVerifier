STYLESHEET = """
/* GLOBAL WINDOW SETTINGS */
QWidget {
    font-family: 'Segoe UI', 'Roboto', sans-serif;
    font-size: 14px;
    background-color: #F3F4F6; /* Light Grey Background */
    color: #333333;
}

/* TABS */
QTabWidget::pane {
    border: 1px solid #E5E7EB;
    background: white;
    border-radius: 6px;
    top: -1px; 
}

QTabBar::tab {
    background: #E5E7EB;
    color: #6B7280;
    padding: 10px 20px;
    margin-right: 4px;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
    font-weight: bold;
}

QTabBar::tab:selected {
    background: white;
    color: #2563EB; /* Bright Blue */
    border-bottom: 2px solid #2563EB;
}

/* TABLE WIDGET */
QTableWidget {
    background-color: white;
    alternate-background-color: #F9FAFB;
    border: 1px solid #D1D5DB;
    border-radius: 6px;
    gridline-color: #E5E7EB;
    selection-background-color: #DBEAFE; /* Light Blue selection */
    selection-color: #1E3A8A;
}

QHeaderView::section {
    background-color: #F3F4F6;
    padding: 6px;
    border: none;
    border-bottom: 2px solid #E5E7EB;
    border-right: 1px solid #E5E7EB;
    font-weight: bold;
    color: #374151;
}

/* BUTTONS */
QPushButton {
    background-color: white;
    border: 1px solid #D1D5DB;
    border-radius: 6px;
    padding: 8px 16px;
    color: #374151;
    font-weight: 600;
}

QPushButton:hover {
    background-color: #F9FAFB;
    border-color: #9CA3AF;
}

QPushButton:pressed {
    background-color: #E5E7EB;
}

/* Primary Action Button (Save) - Blue */
QPushButton#PrimaryBtn {
    background-color: #2563EB;
    color: white;
    border: 1px solid #2563EB;
}
QPushButton#PrimaryBtn:hover {
    background-color: #1D4ED8;
}

/* Success Action Button (Approve) - Green */
QPushButton#SuccessBtn {
    background-color: #10B981;
    color: white;
    border: 1px solid #059669;
}
QPushButton#SuccessBtn:hover {
    background-color: #059669;
}

/* Danger/Warning Button - Red/Orange */
QPushButton#DangerBtn {
    background-color: #EF4444;
    color: white;
    border: 1px solid #DC2626;
}
QPushButton#DangerBtn:hover {
    background-color: #DC2626;
}

/* Info/Reprocess Button - Orange/Yellow */
QPushButton#InfoBtn {
    background-color: #F59E0B;
    color: white;
    border: 1px solid #D97706;
}
QPushButton#InfoBtn:hover {
    background-color: #D97706;
}

/* SPLITTER HANDLE */
QSplitter::handle {
    background-color: #E5E7EB;
    width: 2px;
}

/* IMAGE LABEL */
QLabel#ImageContainer {
    background-color: white;
    border: 2px dashed #D1D5DB;
    border-radius: 8px;
    color: #9CA3AF;
}
QPushButton#DeleteBtn {
    background-color: #FEE2E2;
    color: #DC2626;
    border: 1px solid #FECACA;
    border-radius: 3px;
    font-weight: bold;
    font-size: 16px; /* Slightly larger font fills the center better */
    
    /* CRITICAL FIXES FOR CENTERING */
    padding: 0px;         /* 1. Reset the global padding of 8px 16px */
    padding-bottom: 2px;  /* 2. The "Nudge": Pushes the text visually upwards */
    margin: 0px;
}

QPushButton#DeleteBtn:hover {
    background-color: #EF4444; /* Darker red for better contrast on hover */
    color: white;
    border-color: #DC2626;
}
"""