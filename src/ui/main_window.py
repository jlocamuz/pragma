"""
Ventana principal de la aplicación - Diseño Moderno Corregido
Interfaz gráfica con PyQt5 op
mizada y compatible
"""

import argparse
import sys
import os
from datetime import datetime, timedelta
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QLabel, QDateEdit, QPushButton, QProgressBar, QTextEdit, 
    QGroupBox, QRadioButton, QComboBox, QCheckBox, QMessageBox,
    QFileDialog, QFrame, QGridLayout, QSpacerItem, QSizePolicy, 
    QScrollArea, QDialog, QProgressDialog
)
from PyQt5.QtCore import QDate, QThread, pyqtSignal, Qt, QTimer
from PyQt5.QtGui import QFont, QIcon, QPalette, QColor
from core.data_processor import DataProcessor
from config.default_config import DEFAULT_CONFIG


class ModernCard(QFrame):
    """Card moderno compatible con PyQt5"""
    
    def __init__(self, title="", parent=None):
        super().__init__(parent)
        self.setFrameStyle(QFrame.StyledPanel)
        self.setup_card(title)
    
    def setup_card(self, title):
        """Configura el estilo del card"""
        self.setStyleSheet("""
            ModernCard {
                background-color: white;
                border: 2px solid #e2e8f0;
                border-radius: 8px;
                margin: 4px;
                padding: 8px;
            }
        """)
        
        # Layout principal del card
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 12, 16, 16)
        layout.setSpacing(12)
        
        if title:
            # Título del card
            title_label = QLabel(title)
            title_label.setStyleSheet("""
                QLabel {
                    font-size: 15px;
                    font-weight: bold;
                    color: #1e293b;
                    margin-bottom: 6px;
                }
            """)
            layout.addWidget(title_label)
        
        # Contenedor para el contenido del card
        self.content_layout = QVBoxLayout()
        self.content_layout.setSpacing(8)
        layout.addLayout(self.content_layout)
    
    def add_content(self, widget):
        """Añade contenido al card"""
        self.content_layout.addWidget(widget)


class ModernButton(QPushButton):
    """Botón moderno compatible"""
    
    def __init__(self, text, button_type="primary", parent=None):
        super().__init__(text, parent)
        self.button_type = button_type
        self.setup_button()
    
    def setup_button(self):
        """Configura el estilo del botón"""
        if self.button_type == "primary":
            self.setStyleSheet("""
                QPushButton {
                    background-color: #3b82f6;
                    color: white;
                    border: none;
                    border-radius: 6px;
                    padding: 10px 20px;
                    font-size: 14px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #2563eb;
                }
                QPushButton:pressed {
                    background-color: #1d4ed8;
                }
                QPushButton:disabled {
                    background-color: #94a3b8;
                    color: #e2e8f0;
                }
            """)
        elif self.button_type == "secondary":
            self.setStyleSheet("""
                QPushButton {
                    background-color: #f1f5f9;
                    color: #475569;
                    border: 1px solid #e2e8f0;
                    border-radius: 6px;
                    padding: 10px 20px;
                    font-size: 14px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #e2e8f0;
                    border-color: #cbd5e1;
                }
                QPushButton:pressed {
                    background-color: #cbd5e1;
                }
            """)
        elif self.button_type == "success":
            self.setStyleSheet("""
                QPushButton {
                    background-color: #10b981;
                    color: white;
                    border: none;
                    border-radius: 6px;
                    padding: 10px 20px;
                    font-size: 14px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #059669;
                }
                QPushButton:pressed {
                    background-color: #047857;
                }
            """)


class StatusIndicator(QLabel):
    """Indicador de estado simple"""
    
    def __init__(self, status="idle", text="", parent=None):
        super().__init__(text, parent)
        self.status = status
        self.setup_indicator()
    
    def setup_indicator(self):
        """Configura el indicador"""
        self.setStyleSheet("""
            QLabel {
                padding: 6px 12px;
                border-radius: 12px;
                font-size: 12px;
                font-weight: bold;
            }
        """)
        self.update_status(self.status)
    
    def update_status(self, status, text=""):
        """Actualiza el estado del indicador"""
        self.status = status
        if text:
            self.setText(text)
        
        if status == "success":
            self.setStyleSheet("""
                QLabel {
                    background-color: #dcfce7;
                    color: #166534;
                    border: 1px solid #bbf7d0;
                    padding: 6px 12px;
                    border-radius: 12px;
                    font-size: 12px;
                    font-weight: bold;
                }
            """)
        elif status == "error":
            self.setStyleSheet("""
                QLabel {
                    background-color: #fef2f2;
                    color: #dc2626;
                    border: 1px solid #fecaca;
                    padding: 6px 12px;
                    border-radius: 12px;
                    font-size: 12px;
                    font-weight: bold;
                }
            """)
        elif status == "warning":
            self.setStyleSheet("""
                QLabel {
                    background-color: #fefce8;
                    color: #ca8a04;
                    border: 1px solid #fef08a;
                    padding: 6px 12px;
                    border-radius: 12px;
                    font-size: 12px;
                    font-weight: bold;
                }
            """)
        else:  # idle
            self.setStyleSheet("""
                QLabel {
                    background-color: #f8fafc;
                    color: #64748b;
                    border: 1px solid #e2e8f0;
                    padding: 6px 12px;
                    border-radius: 12px;
                    font-size: 12px;
                    font-weight: bold;
                }
            """)


class InitializationThread(QThread):
    """Thread para la inicialización de la aplicación"""
    progress_updated = pyqtSignal(int, str)
    initialization_finished = pyqtSignal(bool, str, dict)
    
    def __init__(self, processor):
        super().__init__()
        self.processor = processor
    
    def run(self):
        """Ejecuta la inicialización en segundo plano"""
        try:
            # Paso 1: Probar conexión (30%)
            self.progress_updated.emit(10, "🔗 Conectando con la API...")
            success, message = self.processor.test_connection()
            
            if not success:
                self.initialization_finished.emit(False, message, {})
                return
            
            self.progress_updated.emit(30, "✅ Conexión establecida")
            
            # Paso 2: Cargar usuarios (70%)
            self.progress_updated.emit(40, "📋 Cargando usuarios desde API...")
            
            def progress_callback(progress, message):
                mapped_progress = 40 + int((progress / 100) * 50)
                self.progress_updated.emit(mapped_progress, message)
            
            filters = self.processor.get_available_filters(progress_callback)
            
            # Paso 3: Finalizar (100%)
            self.progress_updated.emit(100, "✅ Inicialización completada")
            
            self.initialization_finished.emit(True, "Inicialización exitosa", filters)
            
        except Exception as e:
            error_msg = f"Error durante la inicialización: {str(e)}"
            self.initialization_finished.emit(False, error_msg, {})


class LoadingDialog(QDialog):
    """Diálogo modal de progreso mejorado y legible"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()
    
    def init_ui(self):
        """Inicializa la interfaz del diálogo mejorada"""
        self.setWindowTitle("Cargando Sistema")
        self.setFixedSize(480, 240)
        self.setModal(True)
        self.setWindowFlags(Qt.Dialog | Qt.CustomizeWindowHint | Qt.WindowTitleHint)
        
        # Layout principal con más espacio
        layout = QVBoxLayout(self)
        layout.setSpacing(20)
        layout.setContentsMargins(40, 35, 40, 35)
        
        # Título principal más grande y claro
        title_label = QLabel("🚀 Cargando Sistema")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("""
            QLabel {
                font-size: 22px;
                font-weight: bold;
                color: #1e293b;
                margin-bottom: 10px;
            }
        """)
        layout.addWidget(title_label)
        
        # Subtítulo explicativo
        subtitle_label = QLabel("Configurando reportes de asistencia")
        subtitle_label.setAlignment(Qt.AlignCenter)
        subtitle_label.setStyleSheet("""
            QLabel {
                font-size: 14px;
                color: #64748b;
                margin-bottom: 15px;
            }
        """)
        layout.addWidget(subtitle_label)
        
        # Mensaje de estado más visible
        self.status_label = QLabel("⏳ Preparando...")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("""
            QLabel {
                font-size: 15px;
                color: #374151;
                font-weight: 500;
                margin-bottom: 15px;
                min-height: 20px;
            }
        """)
        layout.addWidget(self.status_label)
        
        # Barra de progreso más grande y visible
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 2px solid #e2e8f0;
                border-radius: 6px;
                background-color: #f8fafc;
                height: 24px;
                text-align: center;
                font-size: 12px;
                font-weight: bold;
                color: #374151;
            }
            QProgressBar::chunk {
                background-color: #3b82f6;
                border-radius: 4px;
            }
        """)
        layout.addWidget(self.progress_bar)
        
        # Contador de progreso más separado
        self.progress_label = QLabel("0%")
        self.progress_label.setAlignment(Qt.AlignCenter)
        self.progress_label.setStyleSheet("""
            QLabel {
                font-size: 14px;
                font-weight: bold;
                color: #2563eb;
                margin-top: 8px;
                min-height: 18px;
            }
        """)
        layout.addWidget(self.progress_label)
        
        # Mensaje de espera
        wait_label = QLabel("Por favor espera...")
        wait_label.setAlignment(Qt.AlignCenter)
        wait_label.setStyleSheet("""
            QLabel {
                font-size: 12px;
                color: #9ca3af;
                font-style: italic;
                margin-top: 5px;
            }
        """)
        layout.addWidget(wait_label)
        
        # Centrar en pantalla
        self.center_on_screen()
        
        # Aplicar estilo general al diálogo mejorado
        self.setStyleSheet("""
            QDialog {
                background-color: white;
                border: 2px solid #d1d5db;
                border-radius: 12px;
            }
        """)
    
    def center_on_screen(self):
        """Centra el diálogo en la pantalla"""
        screen = QApplication.desktop().screenGeometry()
        size = self.geometry()
        self.move(
            (screen.width() - size.width()) // 2,
            (screen.height() - size.height()) // 2
        )
    
    def update_progress(self, progress, message):
        """Actualiza el progreso del diálogo"""
        self.progress_bar.setValue(progress)
        self.progress_label.setText(f"{progress}%")
        self.status_label.setText(message)
        QApplication.processEvents()


class ProcessingThread(QThread):
    """Thread para procesamiento en segundo plano"""
    progress_updated = pyqtSignal(int, str)
    processing_finished = pyqtSignal(dict)
    
    def __init__(self, processor, start_date, end_date, user_ids=None):
        super().__init__()
        self.processor = processor
        self.start_date = start_date
        self.end_date = end_date
        self.user_ids = user_ids
    
    def run(self):
        """Ejecuta el procesamiento en segundo plano"""
        try:
            result = self.processor.process_attendance_report(
                self.start_date, 
                self.end_date, 
                self.user_ids,
                self.progress_callback
            )
            self.processing_finished.emit(result)
        except Exception as e:
            self.processing_finished.emit({
                'success': False,
                'error': str(e),
                'stage': 'thread_error'
            })
    
    def progress_callback(self, progress, message):
        """Callback para actualizar progreso"""
        self.progress_updated.emit(progress, message)


class MainWindow(QMainWindow):
    """Ventana principal moderna de la aplicación"""
    
    def __init__(self):
        super().__init__()
        # NO inicializar processor aquí para evitar carga prematura
        self.processor = None
        self.processing_thread = None
        self.available_users = []
        self.available_filters = {}
        
        self.init_ui()
        # Diferir la inicialización de datos
        QTimer.singleShot(500, self.delayed_initialization)
    
    def delayed_initialization(self):
        """Inicialización diferida para evitar carga prematura"""
        # Crear processor solo cuando sea necesario
        self.processor = DataProcessor()
        self.load_initial_data()
    
    def init_ui(self):
        """Inicializa la interfaz de usuario moderna"""
        self.setWindowTitle("📊 Generador de Reportes de Asistencia - Humand.co")
        self.setGeometry(100, 100, 1000, 800)
        
        # Widget central con scroll
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.setCentralWidget(scroll_area)
        
        # Contenedor principal
        main_widget = QWidget()
        scroll_area.setWidget(main_widget)
        
        # Layout principal
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(16)
        
        # Header
        self.create_header(main_layout)
        
        # Grid de cards
        cards_layout = QGridLayout()
        cards_layout.setSpacing(12)
        
        # Primera fila: Fechas (span completo para más espacio)
        self.create_dates_card(cards_layout, 0, 0, 1, 2)
        
        # Segunda fila: Filtros y Opciones
        self.create_filters_card(cards_layout, 1, 0)
        self.create_options_card(cards_layout, 1, 1)
        
        # Tercera fila: Acciones y Estado (span completo)
        self.create_actions_card(cards_layout, 2, 0, 1, 2)
        self.create_status_card(cards_layout, 3, 0, 1, 2)
        
        main_layout.addLayout(cards_layout)
        
        # Log expandible
        self.create_log_section(main_layout)
        
        # Aplicar estilos globales
        self.apply_global_styles()
    
    def create_header(self, layout):
        """Crea el header moderno"""
        header_widget = QWidget()
        header_layout = QHBoxLayout(header_widget)
        header_layout.setContentsMargins(0, 0, 0, 0)
        
        # Título principal
        title_label = QLabel("📊 Generador de Reportes de Asistencia")
        title_label.setStyleSheet("""
            QLabel {
                font-size: 24px;
                font-weight: bold;
                color: #1e293b;
                margin-bottom: 4px;
            }
        """)
        header_layout.addWidget(title_label)
        
        header_layout.addStretch()
        
        # Estado de conexión en el header
        self.header_status = StatusIndicator("idle", "Desconectado")
        header_layout.addWidget(self.header_status)
        
        layout.addWidget(header_widget)
        
        # Subtítulo
        subtitle_label = QLabel("Sistema automatizado según normativa laboral argentina")
        subtitle_label.setStyleSheet("""
            QLabel {
                font-size: 14px;
                color: #64748b;
                margin-bottom: 12px;
            }
        """)
        layout.addWidget(subtitle_label)
    
    def create_connection_card(self, layout, row, col):
        """Crea el card de conexión"""
        card = ModernCard("🔧 Estado de Conexión")
        
        # Estado de conexión
        self.connection_status = StatusIndicator("idle", "⏳ Esperando inicialización...")
        card.add_content(self.connection_status)
        
        # Botón de prueba
        self.test_connection_btn = ModernButton("🔄 Probar Conexión", "secondary")
        self.test_connection_btn.clicked.connect(self.test_connection)
        self.test_connection_btn.setEnabled(False)  # Deshabilitado hasta inicialización
        card.add_content(self.test_connection_btn)
        
        layout.addWidget(card, row, col)
    
    def create_dates_card(self, layout, row, col, rowspan=1, colspan=1):
        """Crea el card de selección de fechas mejorado y horizontal"""
        card = ModernCard("📅 Selección de Fechas")
        
        # Layout horizontal para las fechas
        dates_layout = QHBoxLayout()
        dates_layout.setSpacing(20)
        
        # Fecha de inicio
        start_container = QVBoxLayout()
        start_label = QLabel("Desde:")
        start_label.setStyleSheet("""
            QLabel {
                font-size: 13px;
                font-weight: bold;
                color: #374151;
                margin-bottom: 4px;
            }
        """)
        start_container.addWidget(start_label)
        
        self.start_date = QDateEdit()
        self.start_date.setDate(QDate.currentDate().addDays(-30))
        self.start_date.setCalendarPopup(True)
        self.start_date.setMinimumHeight(40)
        self.start_date.setStyleSheet("""
            QDateEdit {
                padding: 12px 16px;
                border: 2px solid #e2e8f0;
                border-radius: 8px;
                font-size: 15px;
                font-weight: 500;
                background-color: white;
                min-width: 160px;
                color: #374151;

            }
            QDateEdit:focus {
                border-color: #3b82f6;
                background-color: #f8fafc;
            }
            QDateEdit:hover {
                border-color: #cbd5e1;
            }
        """)
        start_container.addWidget(self.start_date)
        dates_layout.addLayout(start_container)
        
        # Separador visual
        separator = QLabel("→")
        separator.setAlignment(Qt.AlignCenter)
        separator.setStyleSheet("""
            QLabel {
                font-size: 20px;
                color: #94a3b8;
                font-weight: bold;
                margin: 0px 10px;
            }
        """)
        dates_layout.addWidget(separator)
        
        # Fecha de fin
        end_container = QVBoxLayout()
        end_label = QLabel("Hasta:")
        end_label.setStyleSheet(start_label.styleSheet())
        end_container.addWidget(end_label)
        
        self.end_date = QDateEdit()
        self.end_date.setDate(QDate.currentDate())
        self.end_date.setCalendarPopup(True)
        self.end_date.setMinimumHeight(40)
        self.end_date.setStyleSheet(self.start_date.styleSheet())
        end_container.addWidget(self.end_date)
        dates_layout.addLayout(end_container)
        
        # Spacer para empujar el resumen a la derecha
        dates_layout.addStretch()
        
        # Información del rango
        range_container = QVBoxLayout()
        range_label = QLabel("Rango:")
        range_label.setStyleSheet(start_label.styleSheet())
        range_container.addWidget(range_label)
        
        self.range_info_label = QLabel("30 días seleccionados")
        self.range_info_label.setStyleSheet("""
            QLabel {
                font-size: 13px;
                color: #2563eb;
                font-weight: 500;
                background-color: #eff6ff;
                border: 1px solid #dbeafe;
                border-radius: 6px;
                padding: 8px 12px;
                min-width: 120px;
            }
        """)
        range_container.addWidget(self.range_info_label)
        dates_layout.addLayout(range_container)
        
        card.content_layout.addLayout(dates_layout)
        
        # Separador
        separator_line = QFrame()
        separator_line.setFrameShape(QFrame.HLine)
        separator_line.setStyleSheet("QFrame { color: #e2e8f0; margin: 8px 0px; }")
        card.content_layout.addWidget(separator_line)
        
        # Presets de fechas mejorados
        presets_layout = QHBoxLayout()
        presets_layout.setSpacing(8)
        
        # Crear botones de preset más pequeños y elegantes
        preset_buttons = [
            ("Este Mes", 'this_month'),
            ("Mes Anterior", 'last_month'),
            ("Últimos 30 días", 'last_30_days'),
            ("Últimos 7 días", 'last_7_days'),
            ("Esta Semana", 'this_week')
        ]
        
        for text, preset_type in preset_buttons:
            btn = QPushButton(text)
            btn.setStyleSheet("""
                QPushButton {
                    background-color: #f8fafc;
                    color: #475569;
                    border: 1px solid #e2e8f0;
                    border-radius: 6px;
                    padding: 6px 12px;
                    font-size: 12px;
                    font-weight: 500;
                    min-width: 80px;
                }
                QPushButton:hover {
                    background-color: #e2e8f0;
                    border-color: #cbd5e1;
                }
                QPushButton:pressed {
                    background-color: #cbd5e1;
                }
            """)
            btn.clicked.connect(lambda checked, pt=preset_type: self.set_date_preset(pt))
            presets_layout.addWidget(btn)
        
        presets_layout.addStretch()
        card.content_layout.addLayout(presets_layout)
        
        # Conectar eventos para actualizar el rango
        self.start_date.dateChanged.connect(self.update_date_range_info)
        self.end_date.dateChanged.connect(self.update_date_range_info)
        
        # Actualizar información inicial
        self.update_date_range_info()
        
        layout.addWidget(card, row, col, rowspan, colspan)
    
    def create_filters_card(self, layout, row, col):
        """Crea el card de filtros"""
        card = ModernCard("👥 Filtros de Usuarios")
        
        # Opciones de filtrado con contadores
        self.filter_all_users = QRadioButton("Todos los usuarios (cargando...)")
        self.filter_all_users.setChecked(True)
        self.filter_all_users.setStyleSheet("""
            QRadioButton {
                font-size: 14px;
                color: #374151;
                spacing: 8px;
            }
        """)
        card.add_content(self.filter_all_users)
        
        self.filter_by_department = QRadioButton("Por departamento:")
        self.filter_by_department.setStyleSheet(self.filter_all_users.styleSheet())
        card.add_content(self.filter_by_department)
        
        # Layout para combo y contador
        dept_layout = QHBoxLayout()
        self.department_combo = QComboBox()
        self.department_combo.setEnabled(False)
        self.department_combo.setStyleSheet("""
            QComboBox {
                padding: 8px 12px;
                border: 2px solid #e2e8f0;
                border-radius: 6px;
                font-size: 14px;
                background-color: white;
                min-width: 150px;
            }
            QComboBox:focus {
                border-color: #3b82f6;
            }
        """)
        dept_layout.addWidget(self.department_combo)
        
        self.department_count_label = QLabel("(0 empleados)")
        self.department_count_label.setStyleSheet("""
            QLabel {
                color: #64748b;
                font-size: 12px;
                font-weight: bold;
            }
        """)
        dept_layout.addWidget(self.department_count_label)
        dept_layout.addStretch()
        
        card.content_layout.addLayout(dept_layout)
        
        # Conectar eventos
        self.filter_all_users.toggled.connect(self.on_filter_changed)
        self.filter_by_department.toggled.connect(self.on_filter_changed)
        self.department_combo.currentTextChanged.connect(self.update_department_count)
        
        layout.addWidget(card, row, col)
    
    def create_options_card(self, layout, row, col):
        """Crea el card de opciones"""
        card = ModernCard("📊 Opciones de Reporte")
        
        options = [
            ("Resumen consolidado por empleado", True),
            ("Detalle diario", True),
            ("Gráficos y estadísticas", True),
            ("Incluir feriados argentinos", True),
            ("Aplicar compensaciones de horas", True)
        ]
        
        self.option_checkboxes = {}
        
        for option_text, default_checked in options:
            checkbox = QCheckBox(option_text)
            checkbox.setChecked(default_checked)
            checkbox.setStyleSheet("""
                QCheckBox {
                    font-size: 14px;
                    color: #374151;
                    spacing: 8px;
                }
            """)
            card.add_content(checkbox)
            self.option_checkboxes[option_text] = checkbox
        
        layout.addWidget(card, row, col)
    
    def create_actions_card(self, layout, row, col, rowspan=1, colspan=1):
        """Crea el card de acciones"""
        card = ModernCard("🚀 Acciones")
        
        actions_layout = QHBoxLayout()
        
        # Botón principal de generar reporte
        self.generate_report_btn = ModernButton("🚀 GENERAR REPORTE", "primary")
        self.generate_report_btn.clicked.connect(self.generate_report)
        self.generate_report_btn.setEnabled(False)  # Deshabilitado hasta inicialización
        actions_layout.addWidget(self.generate_report_btn)
        
        # Botón secundario
        self.open_folder_btn = ModernButton("📁 Abrir Carpeta", "secondary")
        self.open_folder_btn.clicked.connect(self.open_reports_folder)
        actions_layout.addWidget(self.open_folder_btn)
        
        actions_layout.addStretch()
        
        card.content_layout.addLayout(actions_layout)
        
        layout.addWidget(card, row, col, rowspan, colspan)
    
    def create_status_card(self, layout, row, col, rowspan=1, colspan=1):
        """Crea el card de estado y progreso"""
        card = ModernCard("📈 Estado y Progreso")
        
        # Estado actual
        self.status_label = QLabel("Estado: Inicializando...")
        self.status_label.setStyleSheet("""
            QLabel {
                font-size: 14px;
                color: #374151;
                font-weight: bold;
            }
        """)
        card.add_content(self.status_label)
        
        # Barra de progreso simple
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #e2e8f0;
                border-radius: 4px;
                background-color: #f1f5f9;
                height: 20px;
                text-align: center;
                font-weight: bold;
            }
            QProgressBar::chunk {
                background-color: #10b981;
                border-radius: 3px;
            }
        """)
        card.add_content(self.progress_bar)
        
        # Último reporte
        self.last_report_label = QLabel("Último reporte: Ninguno")
        self.last_report_label.setStyleSheet("""
            QLabel {
                color: #64748b;
                font-size: 12px;
                font-style: italic;
            }
        """)
        card.add_content(self.last_report_label)
        
        layout.addWidget(card, row, col, rowspan, colspan)
    
    def create_log_section(self, layout):
        """Crea la sección de log expandible"""
        log_card = ModernCard("📝 Log de Actividad")
        
        self.log_text = QTextEdit()
        self.log_text.setMaximumHeight(150)
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet("""
            QTextEdit {
                background-color: #f8fafc;
                border: 1px solid #e2e8f0;
                border-radius: 6px;
                font-family: monospace;
                font-size: 11px;
                color: #374151;
                padding: 8px;
            }
        """)
        log_card.add_content(self.log_text)
        
        layout.addWidget(log_card)
    
    def apply_global_styles(self):
        """Aplica estilos globales a la aplicación"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f8fafc;
            }
            QScrollArea {
                border: none;
                background-color: #f8fafc;
            }
            QScrollBar:vertical {
                background-color: #f1f5f9;
                width: 12px;
                border-radius: 6px;
            }
            QScrollBar::handle:vertical {
                background-color: #cbd5e1;
                border-radius: 6px;
                min-height: 20px;
            }
            QScrollBar::handle:vertical:hover {
                background-color: #94a3b8;
            }
            QLabel {
                color: #374151;
            }
        """)
    
    def load_initial_data(self):
        """Carga datos iniciales con popup de progreso"""
        self.log_message("🚀 Iniciando aplicación...")
        QTimer.singleShot(100, self.show_loading_dialog)
    
    def show_loading_dialog(self):
        """Muestra el diálogo de carga usando QProgressDialog nativo"""
        # Crear QProgressDialog nativo - más confiable
        self.progress_dialog = QProgressDialog("Preparando sistema...", None, 0, 100, self)
        self.progress_dialog.setWindowTitle("Cargando Sistema")
        self.progress_dialog.setModal(True)
        self.progress_dialog.setAutoClose(False)
        self.progress_dialog.setAutoReset(False)
        self.progress_dialog.setCancelButton(None)  # Sin botón cancelar
        
        # Configurar estilo del progress dialog
        self.progress_dialog.setStyleSheet("""
            QProgressDialog {
                font-size: 14px;
                min-width: 400px;
                min-height: 120px;
            }
            QProgressBar {
                border: 2px solid #e2e8f0;
                border-radius: 6px;
                background-color: #f8fafc;
                height: 20px;
                text-align: center;
                font-weight: bold;
            }
            QProgressBar::chunk {
                background-color: #3b82f6;
                border-radius: 4px;
            }
        """)
        
        self.progress_dialog.show()
        
        # Iniciar thread
        self.init_thread = InitializationThread(self.processor)
        self.init_thread.progress_updated.connect(self.update_native_progress)
        self.init_thread.initialization_finished.connect(self.initialization_completed)
        self.init_thread.start()
    
    def update_native_progress(self, progress, message):
        """Actualiza el QProgressDialog nativo"""
        if hasattr(self, 'progress_dialog') and self.progress_dialog:
            self.progress_dialog.setValue(progress)
            self.progress_dialog.setLabelText(f"{message}")
            QApplication.processEvents()
    
    def initialization_completed(self, success, message, filters):
        """Maneja la finalización de la inicialización"""
        if hasattr(self, 'progress_dialog') and self.progress_dialog:
            self.progress_dialog.close()
        
        if success:
            # Inicialización exitosa
            self.available_filters = filters
            
            # Actualizar interfaz con datos cargados
            total_users = filters.get('total_users', 0)
            self.filter_all_users.setText(f"Todos los usuarios ({total_users} empleados)")
            
            # Cargar departamentos
            self.department_combo.clear()
            self.department_combo.addItems(filters.get('departments', []))
            
            # Actualizar estados
            self.header_status.update_status("success", "Conectado")
            self.status_label.setText("Estado: Listo para procesar")
            
            # Habilitar controles
            self.generate_report_btn.setEnabled(True)
            
            # Log de éxito
            self.log_message("✅ Aplicación inicializada correctamente")
            self.log_message(f"📋 {total_users} usuarios disponibles")
            
        else:
            # Error en inicialización
            self.header_status.update_status("error", "Error")
            self.status_label.setText("Estado: Error en inicialización")
            
            # Log de error
            self.log_message(f"❌ Error en inicialización: {message}")
            
            # Mostrar mensaje de error
            QMessageBox.critical(
                self, "Error de Inicialización",
                f"No se pudo inicializar la aplicación:\n\n{message}\n\n"
                f"Por favor verifica tu conexión a internet y la configuración de la API."
            )
    
    
    def update_department_count(self):
        """Actualiza el contador de empleados por departamento"""
        if not self.processor:
            return
            
        try:
            current_dept = self.department_combo.currentText()
            if current_dept:
                count = self.processor.get_user_count(department=current_dept)
                self.department_count_label.setText(f"({count} empleados)")
            else:
                self.department_count_label.setText("(0 empleados)")
        except Exception as e:
            self.department_count_label.setText("(error)")
            print(f"Error actualizando contador: {str(e)}")
    
    def set_date_preset(self, preset_type):
        """Establece presets de fechas mejorados"""
        today = QDate.currentDate()
        
        if preset_type == 'this_month':
            start = QDate(today.year(), today.month(), 1)
            end = today
        elif preset_type == 'last_month':
            if today.month() == 1:
                start = QDate(today.year() - 1, 12, 1)
                end = QDate(today.year() - 1, 12, 31)
            else:
                start = QDate(today.year(), today.month() - 1, 1)
                end = QDate(today.year(), today.month(), 1).addDays(-1)
        elif preset_type == 'last_30_days':
            start = today.addDays(-30)
            end = today
        elif preset_type == 'last_7_days':
            start = today.addDays(-7)
            end = today
        elif preset_type == 'this_week':
            # Lunes de esta semana
            days_since_monday = today.dayOfWeek() - 1
            start = today.addDays(-days_since_monday)
            end = today
        
        self.start_date.setDate(start)
        self.end_date.setDate(end)
        self.log_message(f"📅 Fechas establecidas: {start.toString('yyyy-MM-dd')} a {end.toString('yyyy-MM-dd')}")
    
    def update_date_range_info(self):
        """Actualiza la información del rango de fechas"""
        try:
            start_date = self.start_date.date()
            end_date = self.end_date.date()
            
            # Calcular días
            days_diff = start_date.daysTo(end_date) + 1
            
            # Formatear fechas para mostrar
            start_str = start_date.toString("dd MMM")
            end_str = end_date.toString("dd MMM yyyy")
            
            # Actualizar label
            if days_diff == 1:
                self.range_info_label.setText("1 día seleccionado")
            else:
                self.range_info_label.setText(f"{days_diff} días seleccionados")
            
            # Cambiar color según el rango
            if days_diff <= 7:
                color_style = "color: #059669; background-color: #ecfdf5; border-color: #bbf7d0;"
            elif days_diff <= 31:
                color_style = "color: #2563eb; background-color: #eff6ff; border-color: #dbeafe;"
            elif days_diff <= 90:
                color_style = "color: #ca8a04; background-color: #fefce8; border-color: #fef08a;"
            else:
                color_style = "color: #dc2626; background-color: #fef2f2; border-color: #fecaca;"
            
            self.range_info_label.setStyleSheet(f"""
                QLabel {{
                    font-size: 13px;
                    font-weight: 500;
                    border: 1px solid;
                    border-radius: 6px;
                    padding: 8px 12px;
                    min-width: 120px;
                    {color_style}
                }}
            """)
            
        except Exception as e:
            print(f"Error actualizando rango de fechas: {str(e)}")
    
    def on_filter_changed(self):
        """Maneja cambios en los filtros"""
        self.department_combo.setEnabled(self.filter_by_department.isChecked())
        self.update_department_count()
    
    def generate_report(self):
        """Genera el reporte de asistencia"""
        if not self.processor:
            QMessageBox.warning(self, "Error", "Sistema no inicializado correctamente.")
            return
            
        if self.processing_thread and self.processing_thread.isRunning():
            QMessageBox.warning(self, "Procesamiento en curso", 
                              "Ya hay un reporte siendo procesado. Por favor espera a que termine.")
            return
        
        # Validar fechas
        start_date_str = self.start_date.date().toString('yyyy-MM-dd')
        end_date_str = self.end_date.date().toString('yyyy-MM-dd')
        
        validation = self.processor.validate_date_range(start_date_str, end_date_str)
        if not validation['is_valid']:
            QMessageBox.critical(self, "Error en fechas", 
                               f"Fechas inválidas:\n" + "\n".join(validation['errors']))
            return
        
        if validation['warnings']:
            reply = QMessageBox.question(self, "Advertencias", 
                                       f"Se encontraron advertencias:\n" + "\n".join(validation['warnings']) + 
                                       "\n\n¿Deseas continuar?",
                                       QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.No:
                return
        
        # Determinar usuarios a procesar
        user_ids = None
        if self.filter_by_department.isChecked():
            department = self.department_combo.currentText()
            if department:
                user_ids = self.processor.filter_users_by_criteria({'department': department})
                self.log_message(f"👥 Filtrando por departamento: {department}")
        
        # Iniciar procesamiento
        self.log_message(f"🚀 Iniciando generación de reporte: {start_date_str} a {end_date_str}")
        self.start_processing(start_date_str, end_date_str, user_ids)
    
    def start_processing(self, start_date, end_date, user_ids=None):
        """Inicia el procesamiento en segundo plano"""
        self.generate_report_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.status_label.setText("Estado: Procesando...")
        
        self.processing_thread = ProcessingThread(self.processor, start_date, end_date, user_ids)
        self.processing_thread.progress_updated.connect(self.update_progress)
        self.processing_thread.processing_finished.connect(self.processing_completed)
        self.processing_thread.start()
    
    def update_progress(self, progress, message):
        """Actualiza el progreso del procesamiento"""
        self.progress_bar.setValue(progress)
        self.status_label.setText(f"Estado: {message}")
        self.log_message(f"📊 {message} ({progress}%)")
    
    def processing_completed(self, result):
        """Maneja la finalización del procesamiento"""
        self.generate_report_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        if result['success']:
            self.status_label.setText("Estado: ¡Reporte completado!")
            excel_path = result['excel_path']
            
            filename = os.path.basename(excel_path)
            self.last_report_label.setText(f"Último reporte: {filename}")
            
            self.log_message(f"✅ Reporte generado exitosamente: {filename}")
            
            reply = QMessageBox.information(
                self, "¡Reporte Completado!", 
                f"El reporte se ha generado exitosamente.\n\n"
                f"📁 Archivo: {filename}\n"
                f"¿Deseas abrir el archivo?",
                QMessageBox.Yes | QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                self.open_file(excel_path)
        else:
            self.status_label.setText("Estado: Error en procesamiento")
            error_msg = result.get('error', 'Error desconocido')
            stage = result.get('stage', 'unknown')
            
            self.log_message(f"❌ Error en {stage}: {error_msg}")
            
            QMessageBox.critical(
                self, "Error en Procesamiento",
                f"Ocurrió un error durante el procesamiento:\n\n"
                f"Etapa: {stage}\n"
                f"Error: {error_msg}\n\n"
                f"Por favor revisa el log para más detalles."
            )
    
    def open_reports_folder(self):
        """Abre la carpeta de reportes"""
        reports_dir = os.path.expanduser(DEFAULT_CONFIG['output_directory'])
        if os.path.exists(reports_dir):
            if sys.platform == "win32":
                os.startfile(reports_dir)
            elif sys.platform == "darwin":
                os.system(f"open '{reports_dir}'")
            else:
                os.system(f"xdg-open '{reports_dir}'")
            self.log_message(f"📁 Abriendo carpeta: {reports_dir}")
        else:
            QMessageBox.information(self, "Carpeta no encontrada", 
                                  f"La carpeta de reportes no existe aún:\n{reports_dir}\n\n"
                                  f"Se creará automáticamente al generar el primer reporte.")
    
    def open_file(self, filepath):
        """Abre un archivo con la aplicación por defecto"""
        try:
            if sys.platform == "win32":
                os.startfile(filepath)
            elif sys.platform == "darwin":
                os.system(f"open '{filepath}'")
            else:
                os.system(f"xdg-open '{filepath}'")
            self.log_message(f"📄 Abriendo archivo: {os.path.basename(filepath)}")
        except Exception as e:
            self.log_message(f"❌ Error abriendo archivo: {str(e)}")
            QMessageBox.warning(self, "Error", f"No se pudo abrir el archivo:\n{str(e)}")
    
    def log_message(self, message):
        """Añade un mensaje al log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}"
        self.log_text.append(formatted_message)
        
        scrollbar = self.log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
        
        document = self.log_text.document()
        if document.blockCount() > 100:
            cursor = self.log_text.textCursor()
            cursor.movePosition(cursor.MoveOperation.Start)
            cursor.select(cursor.SelectionType.BlockUnderCursor)
            cursor.removeSelectedText()
    
    def wheelEvent(self, event):
        """Maneja el scroll con la rueda del mouse"""
        scroll_area = self.centralWidget()
        if isinstance(scroll_area, QScrollArea):
            scrollbar = scroll_area.verticalScrollBar()
            delta = event.angleDelta().y()
            scroll_speed = 3
            scroll_amount = -delta * scroll_speed // 120
            new_value = scrollbar.value() + scroll_amount
            scrollbar.setValue(new_value)
            event.accept()
        else:
            super().wheelEvent(event)
    
    def closeEvent(self, event):
        """Maneja el cierre de la aplicación"""
        if self.processing_thread and self.processing_thread.isRunning():
            reply = QMessageBox.question(
                self, "Procesamiento en curso",
                "Hay un procesamiento en curso. ¿Deseas cerrar la aplicación de todas formas?",
                QMessageBox.Yes | QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                self.processing_thread.terminate()
                self.processing_thread.wait()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()

def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("--api-key", help="API Key de Humand (sin 'Basic')", required=False)
    return parser.parse_args()


def main():

    print("🔑 API Key en uso:", DEFAULT_CONFIG['api_key'])
    """Función principal para ejecutar la aplicación"""

    app = QApplication(sys.argv)
    app.setApplicationName("Generador de Reportes de Asistencia")
    app.setApplicationVersion("2.0")
    
    # Configurar estilo de la aplicación
    app.setStyle('Fusion')
    
    # Crear y mostrar ventana principal
    window = MainWindow()
    window.show()
    
    # Ejecutar aplicación
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
