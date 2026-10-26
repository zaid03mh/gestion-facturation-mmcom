# -*- coding: utf-8 -*-
"""
MM COM AGENCE - ERP Desktop Application
Application ERP professionnelle pour PME marocaines
Python 3.11+ | PyQt6 | SQLite | ReportLab
"""

import sys
import os
import sqlite3
import hashlib
import secrets
from datetime import datetime, date
from decimal import Decimal, ROUND_HALF_UP

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QLineEdit, QTableWidget, QTableWidgetItem,
    QComboBox, QTextEdit, QSpinBox, QDoubleSpinBox, QDateEdit,
    QDialog, QFormLayout, QMessageBox, QFileDialog, QFrame,
    QStackedWidget, QHeaderView, QGroupBox, QScrollArea, QSplitter,
    QTabWidget, QGridLayout, QSizePolicy, QAbstractItemView
)
from PyQt6.QtCore import Qt, QDate, QSize, pyqtSignal
from PyQt6.QtGui import QFont, QPixmap, QIcon, QPalette, QColor, QPainter, QBrush

# ReportLab pour génération PDF
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT

# ============================================================================
# CONFIGURATION
# ============================================================================
APP_NAME = "MM COM AGENCE"
APP_VERSION = "1.0.0"
DB_NAME = "gestion_mmcom.db"
DOCS_FOLDER = "documents_mmcom"
TVA_RATE = Decimal("0.20")  # 20%
CURRENCY = "DH"

# Couleurs du thème
COLORS = {
    'primary': '#1a73e8',
    'primary_dark': '#1557b0',
    'secondary': '#5f6368',
    'success': '#34a853',
    'warning': '#fbbc04',
    'danger': '#ea4335',
    'background': '#f8f9fa',
    'card': '#ffffff',
    'text': '#202124',
    'text_secondary': '#5f6368',
    'border': '#dadce0'
}

# ============================================================================
# STYLES CSS GLOBAUX
# ============================================================================
GLOBAL_STYLESHEET = """
QMainWindow, QWidget {
    background-color: #f8f9fa;
    font-family: 'Segoe UI', Arial, sans-serif;
}

QLabel {
    color: #202124;
}

QPushButton {
    background-color: #1a73e8;
    color: white;
    border: none;
    padding: 5px 10px;
    border-radius: 4px;
    font-weight: bold;
    font-size: 11px;
    min-width: 60px;
    max-height: 30px;
}

QPushButton:hover {
    background-color: #1557b0;
}

QPushButton:pressed {
    background-color: #0d47a1;
}

QPushButton#secondary {
    background-color: #5f6368;
    color: white;
}

QPushButton#secondary:hover {
    background-color: #4a4e51;
}

QPushButton#success {
    background-color: #34a853;
    color: white;
}

QPushButton#success:hover {
    background-color: #2d9249;
}

QPushButton#danger {
    background-color: #ea4335;
    color: white;
}

QPushButton#danger:hover {
    background-color: #d33426;
}

QPushButton#warning {
    background-color: #fbbc04;
    color: #202124;
}

QLineEdit, QTextEdit, QComboBox, QSpinBox, QDoubleSpinBox, QDateEdit {
    padding: 10px;
    border: 1px solid #dadce0;
    border-radius: 6px;
    background-color: white;
    font-size: 13px;
    color: #202124;
}

QLineEdit:focus, QTextEdit:focus, QComboBox:focus, QSpinBox:focus, QDoubleSpinBox:focus, QDateEdit:focus {
    border: 2px solid #1a73e8;
    color: #202124;
}

QTableWidget {
    background-color: white;
    border: 1px solid #dadce0;
    border-radius: 8px;
    gridline-color: #e8eaed;
    color: #202124;
}

QTableWidget::item {
    padding: 4px;
    color: #202124;
}

QTableWidget::item:selected {
    background-color: #e8f0fe;
    color: #1a73e8;
}

QHeaderView::section {
    background-color: #f1f3f4;
    padding: 10px;
    border: none;
    font-weight: bold;
    color: #5f6368;
}

QScrollBar:vertical {
    border: none;
    background-color: #f1f3f4;
    width: 10px;
    border-radius: 5px;
}

QScrollBar::handle:vertical {
    background-color: #dadce0;
    border-radius: 5px;
    min-height: 20px;
}

QScrollBar::handle:vertical:hover {
    background-color: #bdc1c6;
}

QGroupBox {
    font-weight: bold;
    border: 1px solid #dadce0;
    border-radius: 8px;
    margin-top: 10px;
    padding-top: 10px;
}

QGroupBox::title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 5px;
}

QTabWidget::pane {
    border: 1px solid #dadce0;
    border-radius: 8px;
    background-color: white;
}

QTabBar::tab {
    background-color: #f1f3f4;
    padding: 10px 20px;
    margin-right: 2px;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
}

QTabBar::tab:selected {
    background-color: white;
    color: #1a73e8;
    font-weight: bold;
}
"""

# ============================================================================
# UTILITAIRES SÉCURITÉ
# ============================================================================
def hash_password(password: str) -> str:
    """Hache un mot de passe avec PBKDF2-SHA256"""
    salt = secrets.token_hex(16)
    pwd_hash = hashlib.pbkdf2_hmac('sha256', password.encode(), salt.encode(), 100000)
    return f"{salt}${pwd_hash.hex()}"

def verify_password(password: str, stored_hash: str) -> bool:
    """Vérifie un mot de passe contre son hash"""
    try:
        salt, pwd_hash = stored_hash.split('$')
        new_hash = hashlib.pbkdf2_hmac('sha256', password.encode(), salt.encode(), 100000)
        return new_hash.hex() == pwd_hash
    except:
        return False

# ============================================================================
# BASE DE DONNÉES
# ============================================================================
class Database:
    """Gestionnaire de base de données SQLite"""
    
    def __init__(self, db_path: str):
        self.db_path = db_path
        self.conn = None
        self._connect()
        self._init_tables()
        self._init_default_data()
    
    def _connect(self):
        """Établit la connexion à la base de données"""
        self.conn = sqlite3.connect(self.db_path)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys = ON")
    
    def _init_tables(self):
        """Initialise les tables de la base de données"""
        cursor = self.conn.cursor()
        
        # Table utilisateurs
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                role TEXT NOT NULL DEFAULT 'user',
                actif INTEGER DEFAULT 1,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        # Table configuration entreprise
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS company_settings (
                id INTEGER PRIMARY KEY,
                nom TEXT,
                adresse TEXT,
                ice TEXT,
                rc TEXT,
                identifiant_fiscal TEXT,
                taxe_prof TEXT,
                cnss TEXT,
                telephone TEXT,
                email TEXT,
                logo_path TEXT
            )
        """)
        
        # Migration: Ajouter les colonnes manquantes pour les bases de données existantes
        try:
            columns = [info[1] for info in cursor.execute("PRAGMA table_info(company_settings)").fetchall()]
            new_columns = ['rc', 'identifiant_fiscal', 'taxe_prof', 'cnss']
            for col in new_columns:
                if col not in columns:
                    print(f"Adding missing column: {col}")
                    cursor.execute(f"ALTER TABLE company_settings ADD COLUMN {col} TEXT")
        except Exception as e:
            print(f"Migration error: {e}")
        
        # Table clients
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS clients (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                type TEXT NOT NULL DEFAULT 'particulier',
                nom TEXT NOT NULL,
                prenom TEXT,
                telephone TEXT,
                email TEXT,
                adresse TEXT,
                ice TEXT,
                deleted_at TIMESTAMP,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        # Table services
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS services (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                designation TEXT NOT NULL,
                categorie TEXT,
                prix_unitaire REAL NOT NULL,
                actif INTEGER DEFAULT 1,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        # Table devis
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS devis (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                numero TEXT UNIQUE NOT NULL,
                client_id INTEGER NOT NULL,
                date_devis DATE NOT NULL,
                statut TEXT DEFAULT 'en_cours',
                total_ht REAL DEFAULT 0,
                tva REAL DEFAULT 0,
                total_ttc REAL DEFAULT 0,
                notes TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (client_id) REFERENCES clients(id)
            )
        """)
        
        # Table lignes de devis
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS devis_lignes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                devis_id INTEGER NOT NULL,
                service_id INTEGER,
                designation TEXT NOT NULL,
                quantite INTEGER DEFAULT 1,
                prix_unitaire REAL NOT NULL,
                total REAL NOT NULL,
                FOREIGN KEY (devis_id) REFERENCES devis(id) ON DELETE CASCADE,
                FOREIGN KEY (service_id) REFERENCES services(id)
            )
        """)
        
        # Table bons de livraison
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS bons_livraison (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                numero TEXT UNIQUE NOT NULL,
                client_id INTEGER NOT NULL,
                date_bl DATE NOT NULL,
                notes TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (client_id) REFERENCES clients(id)
            )
        """)
        
        # Table lignes bons de livraison
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS bl_lignes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                bl_id INTEGER NOT NULL,
                designation TEXT NOT NULL,
                quantite INTEGER DEFAULT 1,
                FOREIGN KEY (bl_id) REFERENCES bons_livraison(id) ON DELETE CASCADE
            )
        """)
        
        # Table factures
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS factures (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                numero TEXT UNIQUE NOT NULL,
                client_id INTEGER NOT NULL,
                devis_id INTEGER,
                date_facture DATE NOT NULL,
                statut TEXT DEFAULT 'en_cours',
                total_ht REAL DEFAULT 0,
                tva REAL DEFAULT 0,
                total_ttc REAL DEFAULT 0,
                montant_paye REAL DEFAULT 0,
                notes TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (client_id) REFERENCES clients(id),
                FOREIGN KEY (devis_id) REFERENCES devis(id)
            )
        """)
        
        # Table lignes factures
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS facture_lignes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                facture_id INTEGER NOT NULL,
                service_id INTEGER,
                designation TEXT NOT NULL,
                quantite INTEGER DEFAULT 1,
                prix_unitaire REAL NOT NULL,
                total REAL NOT NULL,
                FOREIGN KEY (facture_id) REFERENCES factures(id) ON DELETE CASCADE,
                FOREIGN KEY (service_id) REFERENCES services(id)
            )
        """)
        
        # Table paiements
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS paiements (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                facture_id INTEGER NOT NULL,
                montant REAL NOT NULL,
                mode TEXT NOT NULL,
                date_paiement DATE NOT NULL,
                notes TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (facture_id) REFERENCES factures(id) ON DELETE CASCADE
            )
        """)
        
        self.conn.commit()
    
    def _init_default_data(self):
        """Initialise les données par défaut"""
        cursor = self.conn.cursor()
        
        # Créer l'admin par défaut s'il n'existe pas
        cursor.execute("SELECT COUNT(*) FROM users WHERE username = 'admin'")
        if cursor.fetchone()[0] == 0:
            admin_hash = hash_password("admin123")
            cursor.execute(
                "INSERT INTO users (username, password_hash, role, actif) VALUES (?, ?, ?, ?)",
                ("admin", admin_hash, "admin", 1)
            )
        
        # Initialiser les paramètres entreprise
        cursor.execute("SELECT COUNT(*) FROM company_settings")
        if cursor.fetchone()[0] == 0:
            cursor.execute(
                "INSERT INTO company_settings (id, nom, adresse, ice, telephone, email) VALUES (?, ?, ?, ?, ?, ?)",
                (1, "MM COM AGENCE", "Casablanca, Maroc", "", "", "")
            )
        
        self.conn.commit()
    
    def execute(self, query: str, params: tuple = ()) -> sqlite3.Cursor:
        """Exécute une requête SQL"""
        return self.conn.execute(query, params)
    
    def commit(self):
        """Valide les changements"""
        self.conn.commit()
    
    def fetchone(self, query: str, params: tuple = ()):
        """Récupère une ligne"""
        return self.conn.execute(query, params).fetchone()
    
    def fetchall(self, query: str, params: tuple = ()):
        """Récupère toutes les lignes"""
        return self.conn.execute(query, params).fetchall()

# ============================================================================
# GÉNÉRATEUR DE NUMÉROS
# ============================================================================
class NumberGenerator:
    """Génère des numéros uniques pour documents"""
    
    @staticmethod
    def generate(db: Database, prefix: str, table: str, column: str) -> str:
        """Génère un numéro au format PREFIX-AAAA-XXXXXX"""
        year = datetime.now().year
        pattern = f"{prefix}-{year}-%"
        
        cursor = db.execute(
            f"SELECT {column} FROM {table} WHERE {column} LIKE ? ORDER BY {column} DESC LIMIT 1",
            (pattern,)
        )
        result = cursor.fetchone()
        
        if result:
            last_num = int(result[0].split('-')[-1])
            new_num = last_num + 1
        else:
            new_num = 1
        
        return f"{prefix}-{year}-{new_num:06d}"

# ============================================================================
# GÉNÉRATEUR PDF
# ============================================================================
class PDFGenerator:
    """Génère les documents PDF professionnels"""
    
    def __init__(self, db: Database):
        self.db = db
        self.docs_folder = DOCS_FOLDER
        os.makedirs(self.docs_folder, exist_ok=True)
    
    def _get_company_info(self) -> dict:
        """Récupère les informations de l'entreprise"""
        row = self.db.fetchone("SELECT * FROM company_settings WHERE id = 1")
        if row:
            return dict(row)
        return {}
    
    def _create_header(self, elements: list, title: str, numero: str, date_doc: str):
        """Crée l'en-tête du document"""
        styles = getSampleStyleSheet()
        company = self._get_company_info()
        
        # Style personnalisé pour le titre
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=24,
            textColor=colors.HexColor('#1a73e8'),
            alignment=TA_CENTER,
            spaceAfter=20
        )
        
        # En-tête entreprise aligné à gauche
        company_style = ParagraphStyle(
            'Company',
            parent=styles['Normal'],
            fontSize=10,
            textColor=colors.HexColor('#5f6368'),
            alignment=TA_LEFT
        )
        
        company_name_style = ParagraphStyle(
            'CompanyName',
            parent=styles['Heading2'],
            alignment=TA_LEFT
        )
        
        # Logo si disponible (aligné à gauche)
        logo_path = company.get('logo_path', '')
        if logo_path and os.path.exists(logo_path):
            try:
                img = Image(logo_path, width=4*cm, height=2*cm)
                img.hAlign = 'LEFT'
                elements.append(img)
            except:
                pass
        
        # Nom entreprise
        elements.append(Paragraph(f"<b>{company.get('nom', APP_NAME)}</b>", company_name_style))
        elements.append(Paragraph(company.get('adresse', ''), company_style))
        if company.get('telephone'):
            elements.append(Paragraph(f"Tél: {company.get('telephone')}", company_style))
        if company.get('email'):
            elements.append(Paragraph(f"Email: {company.get('email')}", company_style))
        
        # Identifiants fiscaux sur une ou plusieurs lignes
        ids = []
        if company.get('ice'): ids.append(f"<b>ICE:</b> {company['ice']}")
        if company.get('rc'): ids.append(f"<b>RC:</b> {company['rc']}")
        if company.get('identifiant_fiscal'): ids.append(f"<b>IF:</b> {company['identifiant_fiscal']}")
        if company.get('taxe_prof'): ids.append(f"<b>TP:</b> {company['taxe_prof']}")
        if company.get('cnss'): ids.append(f"<b>CNSS:</b> {company['cnss']}")
        
        if ids:
            elements.append(Paragraph(" | ".join(ids), company_style))
        
        elements.append(Spacer(1, 20))
        
        # Titre document
        elements.append(Paragraph(title, title_style))
        
        # Numéro et date
        info_style = ParagraphStyle(
            'Info',
            parent=styles['Normal'],
            fontSize=11,
            alignment=TA_CENTER
        )
        elements.append(Paragraph(f"<b>N°:</b> {numero} | <b>Date:</b> {date_doc}", info_style))
        elements.append(Spacer(1, 20))
    
    def _create_client_section(self, elements: list, client: dict):
        """Crée la section client"""
        styles = getSampleStyleSheet()
        
        client_style = ParagraphStyle(
            'Client',
            parent=styles['Normal'],
            fontSize=10,
            leftIndent=20
        )
        
        elements.append(Paragraph("<b>CLIENT</b>", styles['Heading3']))
        elements.append(Paragraph(f"{client.get('nom', '')} {client.get('prenom', '')}", client_style))
        if client.get('adresse'):
            elements.append(Paragraph(client.get('adresse'), client_style))
        if client.get('telephone'):
            elements.append(Paragraph(f"Tél: {client.get('telephone')}", client_style))
        if client.get('ice'):
            elements.append(Paragraph(f"ICE: {client.get('ice')}", client_style))
        
        elements.append(Spacer(1, 20))
    
    def _create_items_table(self, elements: list, items: list, show_prices: bool = True):
        """Crée le tableau des articles"""
        if show_prices:
            headers = ['#', 'Désignation', 'Qté', 'Prix Unit.', 'Total']
            col_widths = [1*cm, 9*cm, 2*cm, 3*cm, 3*cm]
        else:
            headers = ['#', 'Désignation', 'Quantité']
            col_widths = [1*cm, 13*cm, 4*cm]
        
        data = [headers]
        for i, item in enumerate(items, 1):
            if show_prices:
                row = [
                    str(i),
                    item.get('designation', ''),
                    str(item.get('quantite', 1)),
                    f"{item.get('prix_unitaire', 0):.2f} {CURRENCY}",
                    f"{item.get('total', 0):.2f} {CURRENCY}"
                ]
            else:
                row = [
                    str(i),
                    item.get('designation', ''),
                    str(item.get('quantite', 1))
                ]
            data.append(row)
        
        table = Table(data, colWidths=col_widths)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a73e8')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('ALIGN', (0, 1), (0, -1), 'CENTER'),
            ('ALIGN', (1, 1), (1, -1), 'LEFT'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#dadce0')),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f8f9fa')])
        ]))
        
        elements.append(table)
        elements.append(Spacer(1, 20))
    
    def _create_totals_section(self, elements: list, total_ht: float, tva: float, total_ttc: float):
        """Crée la section des totaux"""
        styles = getSampleStyleSheet()
        
        totals_data = [
            ['Total HT:', f"{total_ht:.2f} {CURRENCY}"],
            [f'TVA ({int(TVA_RATE * 100)}%):', f"{tva:.2f} {CURRENCY}"],
            ['Total TTC:', f"{total_ttc:.2f} {CURRENCY}"]
        ]
        
        totals_table = Table(totals_data, colWidths=[12*cm, 6*cm])
        totals_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (0, -1), 'RIGHT'),
            ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, -1), (-1, -1), 12),
            ('TEXTCOLOR', (0, -1), (-1, -1), colors.HexColor('#1a73e8')),
            ('LINEABOVE', (0, -1), (-1, -1), 1, colors.HexColor('#1a73e8'))
        ]))
        
        elements.append(totals_table)
    
    def generate_devis(self, devis_id: int) -> str:
        """Génère le PDF d'un devis"""
        devis = self.db.fetchone("SELECT * FROM devis WHERE id = ?", (devis_id,))
        client = self.db.fetchone("SELECT * FROM clients WHERE id = ?", (devis['client_id'],))
        lignes = self.db.fetchall("SELECT * FROM devis_lignes WHERE devis_id = ?", (devis_id,))
        
        filename = f"{devis['numero']}.pdf"
        filepath = os.path.join(self.docs_folder, filename)
        
        doc = SimpleDocTemplate(filepath, pagesize=A4, topMargin=1*cm, bottomMargin=1*cm)
        elements = []
        
        self._create_header(elements, "DEVIS", devis['numero'], devis['date_devis'])
        self._create_client_section(elements, dict(client))
        self._create_items_table(elements, [dict(l) for l in lignes])
        self._create_totals_section(elements, devis['total_ht'], devis['tva'], devis['total_ttc'])
        
        doc.build(elements)
        return filepath
    
    def generate_bon_livraison(self, bl_id: int) -> str:
        """Génère le PDF d'un bon de livraison"""
        bl = self.db.fetchone("SELECT * FROM bons_livraison WHERE id = ?", (bl_id,))
        client = self.db.fetchone("SELECT * FROM clients WHERE id = ?", (bl['client_id'],))
        lignes = self.db.fetchall("SELECT * FROM bl_lignes WHERE bl_id = ?", (bl_id,))
        
        filename = f"{bl['numero']}.pdf"
        filepath = os.path.join(self.docs_folder, filename)
        
        doc = SimpleDocTemplate(filepath, pagesize=A4, topMargin=1*cm, bottomMargin=1*cm)
        elements = []
        
        self._create_header(elements, "BON DE LIVRAISON", bl['numero'], bl['date_bl'])
        self._create_client_section(elements, dict(client))
        self._create_items_table(elements, [dict(l) for l in lignes], show_prices=False)
        
        doc.build(elements)
        return filepath
    
    def generate_facture(self, facture_id: int) -> str:
        """Génère le PDF d'une facture"""
        facture = self.db.fetchone("SELECT * FROM factures WHERE id = ?", (facture_id,))
        client = self.db.fetchone("SELECT * FROM clients WHERE id = ?", (facture['client_id'],))
        lignes = self.db.fetchall("SELECT * FROM facture_lignes WHERE facture_id = ?", (facture_id,))
        
        filename = f"{facture['numero']}.pdf"
        filepath = os.path.join(self.docs_folder, filename)
        
        doc = SimpleDocTemplate(filepath, pagesize=A4, topMargin=1*cm, bottomMargin=1*cm)
        elements = []
        
        self._create_header(elements, "FACTURE", facture['numero'], facture['date_facture'])
        self._create_client_section(elements, dict(client))
        self._create_items_table(elements, [dict(l) for l in lignes])
        self._create_totals_section(elements, facture['total_ht'], facture['tva'], facture['total_ttc'])
        
        # Ajouter info paiement
        styles = getSampleStyleSheet()
        elements.append(Spacer(1, 20))
        
        montant_paye = facture['montant_paye'] or 0
        reste = facture['total_ttc'] - montant_paye
        
        payment_style = ParagraphStyle('Payment', parent=styles['Normal'], fontSize=10)
        elements.append(Paragraph(f"<b>Montant payé:</b> {montant_paye:.2f} {CURRENCY}", payment_style))
        elements.append(Paragraph(f"<b>Reste à payer:</b> {reste:.2f} {CURRENCY}", payment_style))
        
        doc.build(elements)
        return filepath


# ============================================================================
# WIDGETS PERSONNALISÉS
# ============================================================================
class CardWidget(QFrame):
    """Widget carte avec ombre et style moderne"""
    
    def __init__(self, title: str = "", parent=None):
        super().__init__(parent)
        self.setObjectName("card")
        self.setStyleSheet("""
            QFrame#card {
                background-color: white;
                border-radius: 12px;
                border: 1px solid #e8eaed;
            }
        """)
        
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(20, 20, 20, 20)
        self.layout.setSpacing(15)
        
        if title:
            title_label = QLabel(title)
            title_label.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
            title_label.setStyleSheet("color: #202124;")
            self.layout.addWidget(title_label)
    
    def addWidget(self, widget):
        self.layout.addWidget(widget)
    
    def addLayout(self, layout):
        self.layout.addLayout(layout)


class StatCard(QFrame):
    """Carte de statistique pour le tableau de bord"""
    
    def __init__(self, title: str, value: str, color: str = "#1a73e8", parent=None):
        super().__init__(parent)
        self.setFixedHeight(120)
        self.setStyleSheet(f"""
            QFrame {{
                background-color: white;
                border-radius: 12px;
                border: 1px solid #e8eaed;
                border-left: 4px solid {color};
            }}
        """)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 15, 20, 15)
        
        title_label = QLabel(title)
        title_label.setFont(QFont("Segoe UI", 11))
        title_label.setStyleSheet("color: #5f6368;")
        layout.addWidget(title_label)
        
        self.value_label = QLabel(value)
        self.value_label.setFont(QFont("Segoe UI", 24, QFont.Weight.Bold))
        self.value_label.setStyleSheet(f"color: {color};")
        layout.addWidget(self.value_label)
        layout.addStretch()
    
    def setValue(self, value: str):
        self.value_label.setText(value)


class ActionButton(QPushButton):
    """Bouton d'action stylisé"""
    
    def __init__(self, text: str, style: str = "primary", parent=None):
        super().__init__(text, parent)
        self.setObjectName(style)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setMinimumHeight(40)


# ============================================================================
# DIALOGUE DE CONNEXION
# ============================================================================
class LoginDialog(QDialog):
    """Dialogue de connexion sécurisé"""
    
    def __init__(self, db: Database, parent=None):
        super().__init__(parent)
        self.db = db
        self.user = None
        self.setWindowTitle(f"{APP_NAME} - Connexion")
        self.setFixedSize(400, 350)
        self.setStyleSheet(GLOBAL_STYLESHEET)
        self._setup_ui()
    
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 40, 40, 40)
        layout.setSpacing(20)
        
        title = QLabel(APP_NAME)
        title.setFont(QFont("Segoe UI", 24, QFont.Weight.Bold))
        title.setStyleSheet("color: #1a73e8;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)
        
        subtitle = QLabel("Système de Gestion Intégré")
        subtitle.setFont(QFont("Segoe UI", 11))
        subtitle.setStyleSheet("color: #5f6368;")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(subtitle)
        
        layout.addSpacing(20)
        
        self.username_input = QLineEdit()
        self.username_input.setPlaceholderText("Nom d'utilisateur")
        self.username_input.setMinimumHeight(45)
        layout.addWidget(self.username_input)
        
        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("Mot de passe")
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.password_input.setMinimumHeight(45)
        self.password_input.returnPressed.connect(self._login)
        layout.addWidget(self.password_input)
        
        login_btn = ActionButton("Se connecter")
        login_btn.setMinimumHeight(45)
        login_btn.clicked.connect(self._login)
        layout.addWidget(login_btn)
        
        self.error_label = QLabel("")
        self.error_label.setStyleSheet("color: #ea4335;")
        self.error_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.error_label)
        layout.addStretch()
    
    def _login(self):
        username = self.username_input.text().strip()
        password = self.password_input.text()
        
        if not username or not password:
            self.error_label.setText("Veuillez remplir tous les champs")
            return
        
        user = self.db.fetchone(
            "SELECT * FROM users WHERE username = ? AND actif = 1", (username,)
        )
        
        if user and verify_password(password, user['password_hash']):
            self.user = dict(user)
            self.accept()
        else:
            self.error_label.setText("Identifiants incorrects")
            self.password_input.clear()


# ============================================================================
# MODULE DASHBOARD
# ============================================================================
class DashboardModule(QWidget):
    """Tableau de bord avec statistiques"""
    
    def __init__(self, db: Database, parent=None):
        super().__init__(parent)
        self.db = db
        self._setup_ui()
        self.refresh()
    
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(20)
        
        title = QLabel("Tableau de Bord")
        title.setFont(QFont("Segoe UI", 20, QFont.Weight.Bold))
        layout.addWidget(title)
        
        cards_layout = QHBoxLayout()
        cards_layout.setSpacing(20)
        
        self.clients_card = StatCard("Clients", "0", "#1a73e8")
        self.ca_card = StatCard("CA du Mois", "0 DH", "#34a853")
        self.impaye_card = StatCard("Impayés", "0 DH", "#ea4335")
        self.devis_card = StatCard("Devis du Jour", "0", "#fbbc04")
        
        cards_layout.addWidget(self.clients_card)
        cards_layout.addWidget(self.ca_card)
        cards_layout.addWidget(self.impaye_card)
        cards_layout.addWidget(self.devis_card)
        layout.addLayout(cards_layout)
        
        recent_card = CardWidget("Activité Récente")
        self.recent_table = QTableWidget()
        self.recent_table.setColumnCount(4)
        self.recent_table.setHorizontalHeaderLabels(["Date", "Type", "Numéro", "Client"])
        self.recent_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.recent_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.recent_table.setMinimumHeight(300)
        recent_card.addWidget(self.recent_table)
        layout.addWidget(recent_card)
        layout.addStretch()
    
    def refresh(self):
        result = self.db.fetchone("SELECT COUNT(*) as cnt FROM clients WHERE deleted_at IS NULL")
        self.clients_card.setValue(str(result['cnt']))
        
        first_day = date.today().replace(day=1).isoformat()
        result = self.db.fetchone(
            "SELECT COALESCE(SUM(total_ttc), 0) as total FROM factures WHERE date_facture >= ? AND statut != 'annulee'",
            (first_day,)
        )
        self.ca_card.setValue(f"{result['total']:.2f} DH")
        
        result = self.db.fetchone(
            "SELECT COALESCE(SUM(total_ttc - COALESCE(montant_paye, 0)), 0) as total FROM factures WHERE statut IN ('en_cours', 'partiellement_payee')"
        )
        self.impaye_card.setValue(f"{result['total']:.2f} DH")
        
        today = date.today().isoformat()
        result = self.db.fetchone("SELECT COUNT(*) as cnt FROM devis WHERE date_devis = ?", (today,))
        self.devis_card.setValue(str(result['cnt']))
        
        self._load_recent()
    
    def _load_recent(self):
        self.recent_table.setRowCount(0)
        rows = self.db.fetchall("""
            SELECT date_devis as date, 'Devis' as type, numero,
                   (SELECT nom FROM clients WHERE id = devis.client_id) as client
            FROM devis ORDER BY created_at DESC LIMIT 10
        """)
        for row in rows:
            idx = self.recent_table.rowCount()
            self.recent_table.insertRow(idx)
            self.recent_table.setItem(idx, 0, QTableWidgetItem(str(row['date'])))
            self.recent_table.setItem(idx, 1, QTableWidgetItem(row['type']))
            self.recent_table.setItem(idx, 2, QTableWidgetItem(row['numero']))
            self.recent_table.setItem(idx, 3, QTableWidgetItem(row['client'] or ''))


# ============================================================================
# MODULE CLIENTS
# ============================================================================
class ClientsModule(QWidget):
    """Gestion des clients"""
    
    def __init__(self, db: Database, parent=None):
        super().__init__(parent)
        self.db = db
        self._setup_ui()
        self.refresh()
    
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(20)
        
        header = QHBoxLayout()
        title = QLabel("Gestion des Clients")
        title.setFont(QFont("Segoe UI", 20, QFont.Weight.Bold))
        header.addWidget(title)
        header.addStretch()
        
        add_btn = ActionButton("+ Nouveau Client")
        add_btn.clicked.connect(self._add_client)
        header.addWidget(add_btn)
        layout.addLayout(header)
        
        self.table = QTableWidget()
        self.table.setColumnCount(8)
        self.table.setHorizontalHeaderLabels(["ID", "Type", "Nom", "Téléphone", "Email", "ICE", "Modifier", "Supprimer"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeMode.Fixed)
        self.table.horizontalHeader().setSectionResizeMode(7, QHeaderView.ResizeMode.Fixed)
        self.table.setColumnWidth(6, 100)
        self.table.setColumnWidth(7, 100)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        layout.addWidget(self.table)
    
    def refresh(self):
        self.table.setRowCount(0)
        clients = self.db.fetchall("SELECT * FROM clients WHERE deleted_at IS NULL ORDER BY nom")
        
        for client in clients:
            idx = self.table.rowCount()
            self.table.insertRow(idx)
            self.table.setItem(idx, 0, QTableWidgetItem(str(client['id'])))
            self.table.setItem(idx, 1, QTableWidgetItem(client['type']))
            self.table.setItem(idx, 2, QTableWidgetItem(f"{client['nom']} {client['prenom'] or ''}"))
            self.table.setItem(idx, 3, QTableWidgetItem(client['telephone'] or ''))
            self.table.setItem(idx, 4, QTableWidgetItem(client['email'] or ''))
            self.table.setItem(idx, 5, QTableWidgetItem(client['ice'] or ''))
            
            edit_btn = QPushButton("Modifier")
            edit_btn.setObjectName("secondary")
            edit_btn.clicked.connect(lambda _, c=dict(client): self._edit_client(c))
            self.table.setCellWidget(idx, 6, edit_btn)
            
            del_btn = QPushButton("Supprimer")
            del_btn.setObjectName("danger")
            del_btn.clicked.connect(lambda _, cid=client['id']: self._delete_client(cid))
            self.table.setCellWidget(idx, 7, del_btn)
    
    def _add_client(self):
        dialog = ClientDialog(self.db, parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.refresh()
    
    def _edit_client(self, client: dict):
        dialog = ClientDialog(self.db, client, parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.refresh()
    
    def _delete_client(self, client_id: int):
        reply = QMessageBox.question(
            self, "Confirmer", "Voulez-vous vraiment supprimer ce client?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.db.execute("UPDATE clients SET deleted_at = ? WHERE id = ?", 
                          (datetime.now().isoformat(), client_id))
            self.db.commit()
            self.refresh()


class ClientDialog(QDialog):
    """Dialogue d'ajout/modification de client"""
    
    def __init__(self, db: Database, client: dict = None, parent=None):
        super().__init__(parent)
        self.db = db
        self.client = client
        self.setWindowTitle("Client" if not client else "Modifier Client")
        self.setFixedSize(500, 450)
        self.setStyleSheet(GLOBAL_STYLESHEET)
        self._setup_ui()
    
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(15)
        
        form = QFormLayout()
        form.setSpacing(10)
        
        self.type_combo = QComboBox()
        self.type_combo.addItems(["particulier", "societe"])
        form.addRow("Type:", self.type_combo)
        
        self.nom_input = QLineEdit()
        form.addRow("Nom:", self.nom_input)
        
        self.prenom_input = QLineEdit()
        form.addRow("Prénom:", self.prenom_input)
        
        self.tel_input = QLineEdit()
        form.addRow("Téléphone:", self.tel_input)
        
        self.email_input = QLineEdit()
        form.addRow("Email:", self.email_input)
        
        self.adresse_input = QTextEdit()
        self.adresse_input.setMaximumHeight(80)
        form.addRow("Adresse:", self.adresse_input)
        
        self.ice_input = QLineEdit()
        form.addRow("ICE:", self.ice_input)
        
        layout.addLayout(form)
        
        if self.client:
            self.type_combo.setCurrentText(self.client.get('type', 'particulier'))
            self.nom_input.setText(self.client.get('nom', ''))
            self.prenom_input.setText(self.client.get('prenom', ''))
            self.tel_input.setText(self.client.get('telephone', ''))
            self.email_input.setText(self.client.get('email', ''))
            self.adresse_input.setText(self.client.get('adresse', ''))
            self.ice_input.setText(self.client.get('ice', ''))
        
        buttons = QHBoxLayout()
        save_btn = ActionButton("Enregistrer", "success")
        save_btn.clicked.connect(self._save)
        cancel_btn = ActionButton("Annuler", "secondary")
        cancel_btn.clicked.connect(self.reject)
        buttons.addWidget(cancel_btn)
        buttons.addWidget(save_btn)
        layout.addLayout(buttons)
    
    def _save(self):
        nom = self.nom_input.text().strip()
        if not nom:
            QMessageBox.warning(self, "Erreur", "Le nom est obligatoire")
            return
        
        client_type = self.type_combo.currentText()
        ice = self.ice_input.text().strip()
        
        if client_type == "societe" and not ice:
            QMessageBox.warning(self, "Erreur", "L'ICE est obligatoire pour une société")
            return
        
        data = (
            client_type,
            nom,
            self.prenom_input.text().strip(),
            self.tel_input.text().strip(),
            self.email_input.text().strip(),
            self.adresse_input.toPlainText().strip(),
            ice
        )
        
        if self.client:
            self.db.execute("""
                UPDATE clients SET type=?, nom=?, prenom=?, telephone=?, email=?, adresse=?, ice=?
                WHERE id=?
            """, data + (self.client['id'],))
        else:
            self.db.execute("""
                INSERT INTO clients (type, nom, prenom, telephone, email, adresse, ice)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, data)
        
        self.db.commit()
        self.accept()


# ============================================================================
# MODULE SERVICES
# ============================================================================
class ServicesModule(QWidget):
    """Gestion des services"""
    
    def __init__(self, db: Database, parent=None):
        super().__init__(parent)
        self.db = db
        self._setup_ui()
        self.refresh()
    
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(20)
        
        header = QHBoxLayout()
        title = QLabel("Catalogue des Services")
        title.setFont(QFont("Segoe UI", 20, QFont.Weight.Bold))
        header.addWidget(title)
        header.addStretch()
        
        add_btn = ActionButton("+ Nouveau Service")
        add_btn.clicked.connect(self._add_service)
        header.addWidget(add_btn)
        layout.addLayout(header)
        
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(["ID", "Désignation", "Catégorie", "Prix", "Modifier", "Supprimer"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Fixed)
        self.table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.Fixed)
        self.table.setColumnWidth(4, 100)
        self.table.setColumnWidth(5, 100)
        layout.addWidget(self.table)
    
    def refresh(self):
        self.table.setRowCount(0)
        services = self.db.fetchall("SELECT * FROM services WHERE actif = 1 ORDER BY designation")
        
        for service in services:
            idx = self.table.rowCount()
            self.table.insertRow(idx)
            self.table.setItem(idx, 0, QTableWidgetItem(str(service['id'])))
            self.table.setItem(idx, 1, QTableWidgetItem(service['designation']))
            self.table.setItem(idx, 2, QTableWidgetItem(service['categorie'] or ''))
            self.table.setItem(idx, 3, QTableWidgetItem(f"{service['prix_unitaire']:.2f} DH"))
            
            edit_btn = QPushButton("Modifier")
            edit_btn.setObjectName("secondary")
            edit_btn.clicked.connect(lambda _, s=dict(service): self._edit_service(s))
            self.table.setCellWidget(idx, 4, edit_btn)
            
            del_btn = QPushButton("Supprimer")
            del_btn.setObjectName("danger")
            del_btn.clicked.connect(lambda _, sid=service['id']: self._delete_service(sid))
            self.table.setCellWidget(idx, 5, del_btn)
    
    def _add_service(self):
        dialog = ServiceDialog(self.db, parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.refresh()
    
    def _edit_service(self, service: dict):
        dialog = ServiceDialog(self.db, service, parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.refresh()
    
    def _delete_service(self, service_id: int):
        reply = QMessageBox.question(self, "Confirmer", "Supprimer ce service?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.db.execute("UPDATE services SET actif = 0 WHERE id = ?", (service_id,))
            self.db.commit()
            self.refresh()


class ServiceDialog(QDialog):
    """Dialogue service"""
    
    def __init__(self, db: Database, service: dict = None, parent=None):
        super().__init__(parent)
        self.db = db
        self.service = service
        self.setWindowTitle("Service")
        self.setFixedSize(400, 300)
        self.setStyleSheet(GLOBAL_STYLESHEET)
        self._setup_ui()
    
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        
        form = QFormLayout()
        
        self.designation_input = QLineEdit()
        form.addRow("Désignation:", self.designation_input)
        
        self.categorie_input = QLineEdit()
        form.addRow("Catégorie:", self.categorie_input)
        
        self.prix_input = QDoubleSpinBox()
        self.prix_input.setRange(0, 999999)
        self.prix_input.setDecimals(2)
        self.prix_input.setSuffix(" DH")
        form.addRow("Prix unitaire:", self.prix_input)
        
        layout.addLayout(form)
        
        if self.service:
            self.designation_input.setText(self.service.get('designation', ''))
            self.categorie_input.setText(self.service.get('categorie', ''))
            self.prix_input.setValue(self.service.get('prix_unitaire', 0))
        
        buttons = QHBoxLayout()
        save_btn = ActionButton("Enregistrer", "success")
        save_btn.clicked.connect(self._save)
        cancel_btn = ActionButton("Annuler", "secondary")
        cancel_btn.clicked.connect(self.reject)
        buttons.addWidget(cancel_btn)
        buttons.addWidget(save_btn)
        layout.addLayout(buttons)
    
    def _save(self):
        designation = self.designation_input.text().strip()
        if not designation:
            QMessageBox.warning(self, "Erreur", "La désignation est obligatoire")
            return
        
        data = (designation, self.categorie_input.text().strip(), self.prix_input.value())
        
        if self.service:
            self.db.execute("UPDATE services SET designation=?, categorie=?, prix_unitaire=? WHERE id=?",
                          data + (self.service['id'],))
        else:
            self.db.execute("INSERT INTO services (designation, categorie, prix_unitaire) VALUES (?, ?, ?)", data)
        
        self.db.commit()
        self.accept()


# ============================================================================
# MODULE DEVIS
# ============================================================================
class DevisModule(QWidget):
    """Gestion des devis"""
    
    def __init__(self, db: Database, pdf_gen: PDFGenerator, parent=None):
        super().__init__(parent)
        self.db = db
        self.pdf_gen = pdf_gen
        self._setup_ui()
        self.refresh()
    
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(20)
        
        header = QHBoxLayout()
        title = QLabel("Gestion des Devis")
        title.setFont(QFont("Segoe UI", 20, QFont.Weight.Bold))
        header.addWidget(title)
        header.addStretch()
        
        add_btn = ActionButton("+ Nouveau Devis")
        add_btn.clicked.connect(self._add_devis)
        header.addWidget(add_btn)
        layout.addLayout(header)
        
        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels(["Numéro", "Date", "Client", "Total TTC", "Statut", "PDF", "Facturer"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.Fixed)
        self.table.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeMode.Fixed)
        self.table.setColumnWidth(5, 80)
        self.table.setColumnWidth(6, 100)
        layout.addWidget(self.table)
    
    def refresh(self):
        self.table.setRowCount(0)
        devis_list = self.db.fetchall("""
            SELECT d.*, c.nom as client_nom FROM devis d
            JOIN clients c ON d.client_id = c.id
            ORDER BY d.created_at DESC
        """)
        
        for devis in devis_list:
            idx = self.table.rowCount()
            self.table.insertRow(idx)
            self.table.setItem(idx, 0, QTableWidgetItem(devis['numero']))
            self.table.setItem(idx, 1, QTableWidgetItem(str(devis['date_devis'])))
            self.table.setItem(idx, 2, QTableWidgetItem(devis['client_nom']))
            self.table.setItem(idx, 3, QTableWidgetItem(f"{devis['total_ttc']:.2f} DH"))
            
            statut_item = QTableWidgetItem(devis['statut'])
            if devis['statut'] == 'valide':
                statut_item.setBackground(QColor("#d4edda"))
            elif devis['statut'] == 'facture':
                statut_item.setBackground(QColor("#cce5ff"))
            self.table.setItem(idx, 4, statut_item)
            
            pdf_btn = QPushButton("PDF")
            pdf_btn.setObjectName("secondary")
            pdf_btn.clicked.connect(lambda _, did=devis['id']: self._gen_pdf(did))
            self.table.setCellWidget(idx, 5, pdf_btn)
            
            fac_btn = QPushButton("→ Facture")
            fac_btn.setObjectName("success")
            fac_btn.clicked.connect(lambda _, d=dict(devis): self._convert_to_facture(d))
            self.table.setCellWidget(idx, 6, fac_btn)
    
    def _add_devis(self):
        dialog = DevisDialog(self.db, parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.refresh()
    
    def _gen_pdf(self, devis_id: int):
        filepath = self.pdf_gen.generate_devis(devis_id)
        QMessageBox.information(self, "PDF Généré", f"Fichier créé:\n{filepath}")
        os.startfile(filepath)
    
    def _convert_to_facture(self, devis: dict):
        if devis['statut'] == 'facture':
            QMessageBox.warning(self, "Erreur", "Ce devis a déjà été converti en facture")
            return
        
        reply = QMessageBox.question(self, "Confirmation", "Convertir ce devis en facture?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            numero = NumberGenerator.generate(self.db, "FAC", "factures", "numero")
            
            self.db.execute("""
                INSERT INTO factures (numero, client_id, devis_id, date_facture, total_ht, tva, total_ttc)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (numero, devis['client_id'], devis['id'], date.today().isoformat(),
                  devis['total_ht'], devis['tva'], devis['total_ttc']))
            
            facture_id = self.db.fetchone("SELECT last_insert_rowid() as id")['id']
            
            lignes = self.db.fetchall("SELECT * FROM devis_lignes WHERE devis_id = ?", (devis['id'],))
            for ligne in lignes:
                self.db.execute("""
                    INSERT INTO facture_lignes (facture_id, service_id, designation, quantite, prix_unitaire, total)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (facture_id, ligne['service_id'], ligne['designation'],
                      ligne['quantite'], ligne['prix_unitaire'], ligne['total']))
            
            self.db.execute("UPDATE devis SET statut = 'facture' WHERE id = ?", (devis['id'],))
            self.db.commit()
            
            QMessageBox.information(self, "Succès", f"Facture {numero} créée")
            self.refresh()


class DevisDialog(QDialog):
    """Dialogue création devis"""
    
    def __init__(self, db: Database, parent=None):
        super().__init__(parent)
        self.db = db
        self.lignes = []
        self.setWindowTitle("Nouveau Devis")
        self.setMinimumSize(800, 600)
        self.setStyleSheet(GLOBAL_STYLESHEET)
        self._setup_ui()
    
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        
        form = QHBoxLayout()
        
        self.client_combo = QComboBox()
        clients = self.db.fetchall("SELECT id, nom FROM clients WHERE deleted_at IS NULL ORDER BY nom")
        for c in clients:
            self.client_combo.addItem(c['nom'], c['id'])
        form.addWidget(QLabel("Client:"))
        form.addWidget(self.client_combo)
        
        self.date_edit = QDateEdit()
        self.date_edit.setDate(QDate.currentDate())
        self.date_edit.setCalendarPopup(True)
        form.addWidget(QLabel("Date:"))
        form.addWidget(self.date_edit)
        form.addStretch()
        layout.addLayout(form)
        
        # Lignes
        ligne_group = QGroupBox("Lignes du Devis")
        ligne_layout = QVBoxLayout(ligne_group)
        
        add_ligne_layout = QHBoxLayout()
        self.service_combo = QComboBox()
        self.service_combo.addItem("-- Sélectionner --", None)
        services = self.db.fetchall("SELECT * FROM services WHERE actif = 1 ORDER BY designation")
        for s in services:
            self.service_combo.addItem(f"{s['designation']} - {s['prix_unitaire']:.2f} DH", s['id'])
        self.service_combo.currentIndexChanged.connect(self._on_service_selected)
        add_ligne_layout.addWidget(self.service_combo)
        
        self.designation_input = QLineEdit()
        self.designation_input.setPlaceholderText("Désignation")
        add_ligne_layout.addWidget(self.designation_input)
        
        self.qte_spin = QSpinBox()
        self.qte_spin.setRange(1, 9999)
        self.qte_spin.setValue(1)
        add_ligne_layout.addWidget(QLabel("Qté:"))
        add_ligne_layout.addWidget(self.qte_spin)
        
        self.prix_spin = QDoubleSpinBox()
        self.prix_spin.setRange(0, 999999)
        self.prix_spin.setDecimals(2)
        add_ligne_layout.addWidget(QLabel("Prix:"))
        add_ligne_layout.addWidget(self.prix_spin)
        
        add_btn = ActionButton("Ajouter")
        add_btn.clicked.connect(self._add_ligne)
        add_ligne_layout.addWidget(add_btn)
        ligne_layout.addLayout(add_ligne_layout)
        
        self.lignes_table = QTableWidget()
        self.lignes_table.setColumnCount(5)
        self.lignes_table.setHorizontalHeaderLabels(["Désignation", "Qté", "Prix Unit.", "Total", ""])
        self.lignes_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        ligne_layout.addWidget(self.lignes_table)
        
        layout.addWidget(ligne_group)
        
        # Totaux
        totaux = QHBoxLayout()
        totaux.addStretch()
        
        self.total_ht_label = QLabel("Total HT: 0.00 DH")
        self.total_ht_label.setFont(QFont("Segoe UI", 12))
        totaux.addWidget(self.total_ht_label)
        
        self.tva_label = QLabel("TVA (20%): 0.00 DH")
        self.tva_label.setFont(QFont("Segoe UI", 12))
        totaux.addWidget(self.tva_label)
        
        self.total_ttc_label = QLabel("Total TTC: 0.00 DH")
        self.total_ttc_label.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        self.total_ttc_label.setStyleSheet("color: #1a73e8;")
        totaux.addWidget(self.total_ttc_label)
        
        layout.addLayout(totaux)
        
        # Boutons
        buttons = QHBoxLayout()
        cancel_btn = ActionButton("Annuler", "secondary")
        cancel_btn.clicked.connect(self.reject)
        save_btn = ActionButton("Enregistrer", "success")
        save_btn.clicked.connect(self._save)
        buttons.addStretch()
        buttons.addWidget(cancel_btn)
        buttons.addWidget(save_btn)
        layout.addLayout(buttons)
    
    def _on_service_selected(self, index):
        service_id = self.service_combo.currentData()
        if service_id:
            service = self.db.fetchone("SELECT * FROM services WHERE id = ?", (service_id,))
            self.designation_input.setText(service['designation'])
            self.prix_spin.setValue(service['prix_unitaire'])
    
    def _add_ligne(self):
        designation = self.designation_input.text().strip()
        if not designation:
            return
        
        qte = self.qte_spin.value()
        prix = self.prix_spin.value()
        total = qte * prix
        
        self.lignes.append({
            'service_id': self.service_combo.currentData(),
            'designation': designation,
            'quantite': qte,
            'prix_unitaire': prix,
            'total': total
        })
        
        self._refresh_lignes_table()
        self.designation_input.clear()
        self.service_combo.setCurrentIndex(0)
        self.qte_spin.setValue(1)
        self.prix_spin.setValue(0)
    
    def _refresh_lignes_table(self):
        self.lignes_table.setRowCount(0)
        total_ht = 0
        
        for i, ligne in enumerate(self.lignes):
            idx = self.lignes_table.rowCount()
            self.lignes_table.insertRow(idx)
            self.lignes_table.setItem(idx, 0, QTableWidgetItem(ligne['designation']))
            self.lignes_table.setItem(idx, 1, QTableWidgetItem(str(ligne['quantite'])))
            self.lignes_table.setItem(idx, 2, QTableWidgetItem(f"{ligne['prix_unitaire']:.2f}"))
            self.lignes_table.setItem(idx, 3, QTableWidgetItem(f"{ligne['total']:.2f}"))
            
            del_btn = QPushButton("X")
            del_btn.setObjectName("danger")
            del_btn.clicked.connect(lambda _, li=i: self._remove_ligne(li))
            self.lignes_table.setCellWidget(idx, 4, del_btn)
            
            total_ht += ligne['total']
        
        tva = total_ht * float(TVA_RATE)
        total_ttc = total_ht + tva
        
        self.total_ht_label.setText(f"Total HT: {total_ht:.2f} DH")
        self.tva_label.setText(f"TVA (20%): {tva:.2f} DH")
        self.total_ttc_label.setText(f"Total TTC: {total_ttc:.2f} DH")
    
    def _remove_ligne(self, index):
        self.lignes.pop(index)
        self._refresh_lignes_table()
    
    def _save(self):
        if not self.lignes:
            QMessageBox.warning(self, "Erreur", "Ajoutez au moins une ligne")
            return
        
        client_id = self.client_combo.currentData()
        if not client_id:
            QMessageBox.warning(self, "Erreur", "Sélectionnez un client")
            return
        
        total_ht = sum(l['total'] for l in self.lignes)
        tva = total_ht * float(TVA_RATE)
        total_ttc = total_ht + tva
        
        numero = NumberGenerator.generate(self.db, "DEV", "devis", "numero")
        date_devis = self.date_edit.date().toString("yyyy-MM-dd")
        
        self.db.execute("""
            INSERT INTO devis (numero, client_id, date_devis, total_ht, tva, total_ttc)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (numero, client_id, date_devis, total_ht, tva, total_ttc))
        
        devis_id = self.db.fetchone("SELECT last_insert_rowid() as id")['id']
        
        for ligne in self.lignes:
            self.db.execute("""
                INSERT INTO devis_lignes (devis_id, service_id, designation, quantite, prix_unitaire, total)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (devis_id, ligne['service_id'], ligne['designation'],
                  ligne['quantite'], ligne['prix_unitaire'], ligne['total']))
        
        self.db.commit()
        self.accept()


# ============================================================================
# MODULE FACTURES
# ============================================================================
class FacturesModule(QWidget):
    """Gestion des factures"""
    
    def __init__(self, db: Database, pdf_gen: PDFGenerator, parent=None):
        super().__init__(parent)
        self.db = db
        self.pdf_gen = pdf_gen
        self._setup_ui()
        self.refresh()
    
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        
        header = QHBoxLayout()
        title = QLabel("Gestion des Factures")
        title.setFont(QFont("Segoe UI", 20, QFont.Weight.Bold))
        header.addWidget(title)
        header.addStretch()
        layout.addLayout(header)
        
        self.table = QTableWidget()
        self.table.setColumnCount(8)
        self.table.setHorizontalHeaderLabels(["Numéro", "Date", "Client", "Total TTC", "Payé", "Statut", "PDF", "Paiement"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeMode.Fixed)
        self.table.horizontalHeader().setSectionResizeMode(7, QHeaderView.ResizeMode.Fixed)
        self.table.setColumnWidth(6, 80)
        self.table.setColumnWidth(7, 100)
        layout.addWidget(self.table)
    
    def refresh(self):
        self.table.setRowCount(0)
        factures = self.db.fetchall("""
            SELECT f.*, c.nom as client_nom FROM factures f
            JOIN clients c ON f.client_id = c.id
            ORDER BY f.created_at DESC
        """)
        
        for fac in factures:
            idx = self.table.rowCount()
            self.table.insertRow(idx)
            self.table.setItem(idx, 0, QTableWidgetItem(fac['numero']))
            self.table.setItem(idx, 1, QTableWidgetItem(str(fac['date_facture'])))
            self.table.setItem(idx, 2, QTableWidgetItem(fac['client_nom']))
            self.table.setItem(idx, 3, QTableWidgetItem(f"{fac['total_ttc']:.2f} DH"))
            self.table.setItem(idx, 4, QTableWidgetItem(f"{fac['montant_paye'] or 0:.2f} DH"))
            
            statut_item = QTableWidgetItem(fac['statut'])
            if fac['statut'] == 'payee':
                statut_item.setBackground(QColor("#d4edda"))
            elif fac['statut'] == 'partiellement_payee':
                statut_item.setBackground(QColor("#fff3cd"))
            self.table.setItem(idx, 5, statut_item)
            
            pdf_btn = QPushButton("PDF")
            pdf_btn.setObjectName("secondary")
            pdf_btn.clicked.connect(lambda _, fid=fac['id']: self._gen_pdf(fid))
            self.table.setCellWidget(idx, 6, pdf_btn)
            
            pay_btn = QPushButton("Paiement")
            pay_btn.setObjectName("success")
            pay_btn.clicked.connect(lambda _, f=dict(fac): self._add_paiement(f))
            self.table.setCellWidget(idx, 7, pay_btn)
    
    def _gen_pdf(self, facture_id: int):
        filepath = self.pdf_gen.generate_facture(facture_id)
        QMessageBox.information(self, "PDF", f"Fichier:\n{filepath}")
        os.startfile(filepath)
    
    def _add_paiement(self, facture: dict):
        dialog = PaiementDialog(self.db, facture, parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.refresh()


class PaiementDialog(QDialog):
    """Dialogue paiement"""
    
    def __init__(self, db: Database, facture: dict, parent=None):
        super().__init__(parent)
        self.db = db
        self.facture = facture
        self.setWindowTitle(f"Paiement - {facture['numero']}")
        self.setFixedSize(400, 350)
        self.setStyleSheet(GLOBAL_STYLESHEET)
        self._setup_ui()
    
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        
        reste = self.facture['total_ttc'] - (self.facture['montant_paye'] or 0)
        
        info = QLabel(f"Total: {self.facture['total_ttc']:.2f} DH\nPayé: {self.facture['montant_paye'] or 0:.2f} DH\nReste: {reste:.2f} DH")
        info.setFont(QFont("Segoe UI", 11))
        layout.addWidget(info)
        
        form = QFormLayout()
        
        self.montant_spin = QDoubleSpinBox()
        self.montant_spin.setRange(0.01, reste)
        self.montant_spin.setValue(reste)
        self.montant_spin.setDecimals(2)
        self.montant_spin.setSuffix(" DH")
        form.addRow("Montant:", self.montant_spin)
        
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["Espèces", "Chèque", "Virement", "Carte"])
        form.addRow("Mode:", self.mode_combo)
        
        self.date_edit = QDateEdit()
        self.date_edit.setDate(QDate.currentDate())
        self.date_edit.setCalendarPopup(True)
        form.addRow("Date:", self.date_edit)
        
        self.notes_input = QLineEdit()
        form.addRow("Notes:", self.notes_input)
        
        layout.addLayout(form)
        
        buttons = QHBoxLayout()
        cancel_btn = ActionButton("Annuler", "secondary")
        cancel_btn.clicked.connect(self.reject)
        save_btn = ActionButton("Enregistrer", "success")
        save_btn.clicked.connect(self._save)
        buttons.addWidget(cancel_btn)
        buttons.addWidget(save_btn)
        layout.addLayout(buttons)
    
    def _save(self):
        montant = self.montant_spin.value()
        mode = self.mode_combo.currentText()
        date_p = self.date_edit.date().toString("yyyy-MM-dd")
        
        self.db.execute("""
            INSERT INTO paiements (facture_id, montant, mode, date_paiement, notes)
            VALUES (?, ?, ?, ?, ?)
        """, (self.facture['id'], montant, mode, date_p, self.notes_input.text()))
        
        nouveau_total = (self.facture['montant_paye'] or 0) + montant
        
        if nouveau_total >= self.facture['total_ttc']:
            statut = 'payee'
        else:
            statut = 'partiellement_payee'
        
        self.db.execute("UPDATE factures SET montant_paye = ?, statut = ? WHERE id = ?",
                       (nouveau_total, statut, self.facture['id']))
        self.db.commit()
        self.accept()


# ============================================================================
# MODULE CONFIGURATION
# ============================================================================
class ConfigModule(QWidget):
    """Configuration entreprise"""
    
    def __init__(self, db: Database, parent=None):
        super().__init__(parent)
        self.db = db
        self._setup_ui()
        self._load()
    
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        
        title = QLabel("Configuration Entreprise")
        title.setFont(QFont("Segoe UI", 20, QFont.Weight.Bold))
        layout.addWidget(title)
        
        card = CardWidget()
        form = QFormLayout()
        
        self.nom_input = QLineEdit()
        form.addRow("Nom société:", self.nom_input)
        
        self.adresse_input = QTextEdit()
        self.adresse_input.setMaximumHeight(80)
        form.addRow("Adresse:", self.adresse_input)
        
        self.ice_input = QLineEdit()
        form.addRow("ICE:", self.ice_input)

        self.rc_input = QLineEdit()
        form.addRow("RC (Registre Commerce):", self.rc_input)

        self.if_input = QLineEdit()
        form.addRow("Identifiant Fiscal:", self.if_input)

        self.tp_input = QLineEdit()
        form.addRow("Taxe Professionnelle:", self.tp_input)

        self.cnss_input = QLineEdit()
        form.addRow("CNSS:", self.cnss_input)
        
        self.tel_input = QLineEdit()
        form.addRow("Téléphone:", self.tel_input)
        
        self.email_input = QLineEdit()
        form.addRow("Email:", self.email_input)
        
        logo_layout = QHBoxLayout()
        self.logo_label = QLabel("Aucun logo")
        self.logo_path = ""
        logo_btn = ActionButton("Choisir Logo", "secondary")
        logo_btn.clicked.connect(self._choose_logo)
        logo_layout.addWidget(self.logo_label)
        logo_layout.addWidget(logo_btn)
        form.addRow("Logo:", logo_layout)
        
        card.addLayout(form)
        
        save_btn = ActionButton("Enregistrer", "success")
        save_btn.clicked.connect(self._save)
        card.addWidget(save_btn)
        
        layout.addWidget(card)
        layout.addStretch()
    
    def _load(self):
        settings = self.db.fetchone("SELECT * FROM company_settings WHERE id = 1")
        if settings:
            self.nom_input.setText(settings['nom'] or '')
            self.adresse_input.setText(settings['adresse'] or '')
            self.ice_input.setText(settings['ice'] or '')
            self.rc_input.setText(settings['rc'] or '')
            self.if_input.setText(settings['identifiant_fiscal'] or '')
            self.tp_input.setText(settings['taxe_prof'] or '')
            self.cnss_input.setText(settings['cnss'] or '')
            self.tel_input.setText(settings['telephone'] or '')
            self.email_input.setText(settings['email'] or '')
            if settings['logo_path']:
                self.logo_path = settings['logo_path']
                self.logo_label.setText(os.path.basename(self.logo_path))
    
    def _choose_logo(self):
        path, _ = QFileDialog.getOpenFileName(self, "Choisir Logo", "", "Images (*.png *.jpg *.jpeg)")
        if path:
            self.logo_path = path
            self.logo_label.setText(os.path.basename(path))
    
    def _save(self):
        self.db.execute("""
            UPDATE company_settings SET nom=?, adresse=?, ice=?, rc=?, identifiant_fiscal=?, taxe_prof=?, cnss=?, telephone=?, email=?, logo_path=?
            WHERE id = 1
        """, (
            self.nom_input.text().strip(),
            self.adresse_input.toPlainText().strip(),
            self.ice_input.text().strip(),
            self.rc_input.text().strip(),
            self.if_input.text().strip(),
            self.tp_input.text().strip(),
            self.cnss_input.text().strip(),
            self.tel_input.text().strip(),
            self.email_input.text().strip(),
            self.logo_path
        ))
        self.db.commit()
        QMessageBox.information(self, "Succès", "Configuration enregistrée")


# ============================================================================
# FENÊTRE PRINCIPALE
# ============================================================================
class MainWindow(QMainWindow):
    """Fenêtre principale de l'application"""
    
    def __init__(self, db: Database, user: dict):
        super().__init__()
        self.db = db
        self.user = user
        self.pdf_gen = PDFGenerator(db)
        
        self.setWindowTitle(f"{APP_NAME} - {user['username']}")
        self.setMinimumSize(1200, 800)
        self.setStyleSheet(GLOBAL_STYLESHEET)
        
        self._setup_ui()
    
    def _setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QHBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # Menu latéral
        sidebar = QFrame()
        sidebar.setFixedWidth(220)
        sidebar.setStyleSheet("""
            QFrame {
                background-color: #1a73e8;
            }
            QPushButton {
                background-color: transparent;
                color: white;
                text-align: left;
                padding: 15px 20px;
                border: none;
                border-radius: 0;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: rgba(255, 255, 255, 0.1);
            }
            QPushButton:checked {
                background-color: rgba(255, 255, 255, 0.2);
                font-weight: bold;
            }
        """)
        
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(0, 20, 0, 20)
        sidebar_layout.setSpacing(5)
        
        # Logo
        logo = QLabel(APP_NAME)
        logo.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        logo.setStyleSheet("color: white; padding: 20px;")
        logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        sidebar_layout.addWidget(logo)
        
        # Boutons menu
        self.menu_buttons = []
        menus = [
            ("📊 Tableau de Bord", 0),
            ("👥 Clients", 1),
            ("📋 Services", 2),
            ("📝 Devis", 3),
            ("📄 Factures", 4),
        ]
        
        if self.user['role'] == 'admin':
            menus.append(("⚙️ Configuration", 5))
        
        for text, index in menus:
            btn = QPushButton(text)
            btn.setCheckable(True)
            btn.clicked.connect(lambda _, i=index: self._switch_page(i))
            self.menu_buttons.append(btn)
            sidebar_layout.addWidget(btn)
        
        sidebar_layout.addStretch()
        
        # Déconnexion
        logout_btn = QPushButton("🚪 Déconnexion")
        logout_btn.clicked.connect(self._logout)
        sidebar_layout.addWidget(logout_btn)
        
        main_layout.addWidget(sidebar)
        
        # Zone de contenu
        self.stack = QStackedWidget()
        
        self.dashboard = DashboardModule(self.db)
        self.clients = ClientsModule(self.db)
        self.services = ServicesModule(self.db)
        self.devis = DevisModule(self.db, self.pdf_gen)
        self.factures = FacturesModule(self.db, self.pdf_gen)
        self.config = ConfigModule(self.db)
        
        self.stack.addWidget(self.dashboard)
        self.stack.addWidget(self.clients)
        self.stack.addWidget(self.services)
        self.stack.addWidget(self.devis)
        self.stack.addWidget(self.factures)
        self.stack.addWidget(self.config)
        
        main_layout.addWidget(self.stack)
        
        self._switch_page(0)
    
    def _switch_page(self, index: int):
        for i, btn in enumerate(self.menu_buttons):
            btn.setChecked(i == index)
        self.stack.setCurrentIndex(index)
        
        # Rafraîchir le module actif
        widget = self.stack.currentWidget()
        if hasattr(widget, 'refresh'):
            widget.refresh()
    
    def _logout(self):
        self.close()
        # Relancer le login
        login = LoginDialog(self.db)
        if login.exec() == QDialog.DialogCode.Accepted:
            window = MainWindow(self.db, login.user)
            window.show()


# ============================================================================
# POINT D'ENTRÉE
# ============================================================================
def main():
    """Point d'entrée de l'application"""
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    # Créer le dossier documents
    os.makedirs(DOCS_FOLDER, exist_ok=True)
    
    # Initialiser la base de données
    db = Database(DB_NAME)
    
    # Afficher le dialogue de connexion
    login = LoginDialog(db)
    if login.exec() == QDialog.DialogCode.Accepted:
        window = MainWindow(db, login.user)
        window.show()
        sys.exit(app.exec())
    else:
        sys.exit(0)


if __name__ == "__main__":
    main()
