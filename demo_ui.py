"""
Demo script to showcase the modern UI improvements
"""
import sys
import os
from PySide6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QLabel, QPushButton
from PySide6.QtCore import Qt

# Add src to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from styles import get_style_sheet

class DemoWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("💼 Modern UI Demo - Salary Calculator")
        self.setGeometry(100, 100, 800, 600)
        
        # Apply modern styling
        self.setStyleSheet(get_style_sheet())
        
        # Create central widget
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setSpacing(20)
        layout.setContentsMargins(30, 30, 30, 30)
        
        # Header
        header = QLabel("🎨 Modern UI Features")
        header.setProperty("class", "header")
        header.setAlignment(Qt.AlignCenter)
        layout.addWidget(header)
        
        # Feature list
        features = [
            "✨ Modern color palette with professional blue theme",
            "🎯 Improved visual hierarchy and spacing",
            "📱 Responsive design with scroll areas",
            "🔘 Enhanced buttons with hover effects",
            "📝 Better form layouts with placeholders",
            "📊 Modern table styling with alternating rows",
            "🚀 Progress indicators and status messages",
            "🎪 Card-based design with rounded corners",
            "🌟 Professional typography and icons"
        ]
        
        for feature in features:
            label = QLabel(feature)
            label.setStyleSheet("padding: 8px; background-color: #f8fafc; border-radius: 6px; margin: 2px;")
            layout.addWidget(label)
        
        # Demo buttons
        btn_layout = QVBoxLayout()
        
        primary_btn = QPushButton("🚀 Primary Action")
        primary_btn.setProperty("class", "success")
        primary_btn.setMinimumHeight(45)
        btn_layout.addWidget(primary_btn)
        
        secondary_btn = QPushButton("📂 Secondary Action")
        secondary_btn.setProperty("class", "secondary")
        secondary_btn.setMinimumHeight(45)
        btn_layout.addWidget(secondary_btn)
        
        danger_btn = QPushButton("⚠️ Danger Action")
        danger_btn.setProperty("class", "danger")
        danger_btn.setMinimumHeight(45)
        btn_layout.addWidget(danger_btn)
        
        layout.addLayout(btn_layout)
        layout.addStretch()
        
        # Status message
        status = QLabel("✨ UI enhancements completed successfully!")
        status.setProperty("class", "status-success")
        status.setAlignment(Qt.AlignCenter)
        layout.addWidget(status)

def main():
    app = QApplication(sys.argv)
    window = DemoWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
