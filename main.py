from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon, QPixmap, QFont
from PyQt5.QtWidgets import QScrollArea
import sys
import time
import os
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel,
    QProgressBar, QHBoxLayout, QSpacerItem, QScrollArea, QSizePolicy,
    QDialog, QTextEdit, QDialogButtonBox, QMessageBox, QGridLayout, QDesktopWidget
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QFont, QPixmap, QIcon


def recurso_caminho(rel_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, rel_path)


class LoaderDialog(QDialog):
    def __init__(self, parent=None, light_theme=True):
        super().__init__(parent)
        self.setWindowTitle("Salvando Arquivo")
        self.resize(400, 120)
        self.setMinimumSize(300, 100)
        layout = QVBoxLayout()
        self.setLayout(layout)

        self.label = QLabel("Carregando arquivo, aguarde...")
        self.label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label)

        self.progressBar = QProgressBar()
        self.progressBar.setMaximum(100)
        self.progressBar.setValue(0)
        layout.addWidget(self.progressBar)

        self.setWindowFlags(self.windowFlags() & ~
                            Qt.WindowContextHelpButtonHint)
        self.setModal(True)

        self.aplicar_tema(light_theme)

    def aplicar_tema(self, light_theme):
        if light_theme:
            self.label.setStyleSheet("""
                color: #000000;
                background-color: #ffffff;
                font-weight: bold;
                border: none;
            """)
            self.progressBar.setStyleSheet("""
                QProgressBar {
                    border: 2px solid #ff6600;
                    border-radius: 10px;
                    background-color: #ffffff;
                    height: 25px;
                    text-align: center;
                    font-weight: bold;
                    color: #000000;
                }
                QProgressBar::chunk {
                    background-color: #ff6600;
                    width: 20px;
                }
            """)
            self.setStyleSheet("background-color: #ffffff;")
        else:
            self.label.setStyleSheet("""
                color: #ffa500;
                background-color: #000000;
                font-weight: bold;
                border: none;
            """)
            self.progressBar.setStyleSheet("""
                QProgressBar {
                    border: 2px solid #ffa500;
                    border-radius: 10px;
                    background-color: #222222;
                    height: 25px;
                    text-align: center;
                    font-weight: bold;
                    color: #ffffff;
                }
                QProgressBar::chunk {
                    background-color: #ffa500;
                    width: 20px;
                }
            """)
            self.setStyleSheet("background-color: #000000;")

    def atualizar_progresso(self, valor):
        self.progressBar.setValue(valor)


class LoaderThread(QThread):
    progress = pyqtSignal(int)

    def run(self):
        for i in range(101):
            self.progress.emit(i)
            time.sleep(0.01)


# ...


class BancoApp(QWidget):
    def __init__(self):
        super().__init__()
        self.light_theme = True  # tema inicial
        # Offset proporcional (percentual da altura da janela)
        self.offset_factor_light = 0.05   # claro: ~5%  (ajuste a gosto)
        # escuro: ~2% (mais alto = mais espaço no topo)
        self.offset_factor_dark = 0.05
        self.background_label = QLabel(self)
        self.background_label.setScaledContents(True)
        self.background_label.lower()
        self.setWindowTitle("Central de Bancos")
        self.resize(983, 750)  # Tamanho inicial
        self.setMinimumSize(700, 500)
        self.setWindowIcon(
            QIcon(recurso_caminho("imagens/icone_principal.ico")))

        # 🔷 ScrollArea e conteúdo interno
        scroll_area = QScrollArea(self)
        scroll_area.setWidgetResizable(True)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        content_widget = QWidget()
        self.img_light = recurso_caminho("imagens/fundo_claro.jpg")
        self.img_dark = recurso_caminho(
            "imagens/fundo_escuro.jpg")  # crie esse arquivo
        self.background_label = QLabel(content_widget)
        self.background_label.setPixmap(
            QPixmap(self.img_light))  # começa no claro
        self.background_label.setScaledContents(True)
        self.background_label.lower()
        self.background_label.setScaledContents(True)
        self.background_label.lower()
        # Layout principal do content_widget
        layout_central = QVBoxLayout(content_widget)

# Layout fixo para o conteúdo real (título, botões, etc.)
        self.layout = QVBoxLayout()
        self.layout.setContentsMargins(30, 30, 30, 20)
        self.layout.setSpacing(15)

# Container com tudo
        conteudo = QWidget()
        conteudo.setLayout(self.layout)
        conteudo.setAttribute(Qt.WA_TranslucentBackground)
        conteudo.setStyleSheet("background: transparent;")

        # Offset superior para descer a grade
        from PyQt5.QtWidgets import QSpacerItem
        self.top_offset = QSpacerItem(
            0, 0, QSizePolicy.Minimum, QSizePolicy.Fixed)  # ajuste 60 à vontade
        self.layout.addItem(self.top_offset)


# Espaço fixo (não expansível) acima/abaixo do conteúdo
        layout_central.addSpacing(0)
        layout_central.addWidget(conteudo, alignment=Qt.AlignTop)
        layout_central.addSpacing(0)
        self.layout.setContentsMargins(30, 30, 30, 20)
        self.layout.setSpacing(15)

        outer_layout = QVBoxLayout(self)
        outer_layout.addWidget(scroll_area)
        scroll_area.setWidget(content_widget)

        self.titulo = QLabel()   # título vazio
        self.titulo.hide()       # não aparece

        self.toggle_btn = QPushButton()
        self.toggle_btn.setFixedSize(130, 34)
        self.toggle_btn.clicked.connect(self.toggle_theme)

        # Barra superior só com o botão de tema (alinhado à direita)
        self.top_bar = QHBoxLayout()
        self.top_bar.setContentsMargins(0, 0, 0, 0)
        self.top_bar.addStretch(1)
        self.top_bar.addWidget(self.toggle_btn, alignment=Qt.AlignRight)

        # Insere a barra no topo do layout da página
        self.layout.insertLayout(0, self.top_bar)

        # 🔹 Navegação por páginas (setinhas)
        self.nav_prev = QPushButton("<<")
        self.nav_next = QPushButton(">>")
        self.nav_prev.setFixedSize(36, 36)   # antes: 40x40
        self.nav_next.setFixedSize(36, 36)
        self.nav_prev.setCursor(Qt.PointingHandCursor)
        self.nav_next.setCursor(Qt.PointingHandCursor)
        self.nav_prev.clicked.connect(self.prev_page)
        self.nav_next.clicked.connect(self.next_page)

        self.page_label = QLabel("")
        self.page_label.setAlignment(Qt.AlignCenter)
        self.page_label.setMinimumWidth(120)
        self.page_label.setStyleSheet("background: transparent;")

        self.nav_bar = QHBoxLayout()
        self.nav_bar.setAlignment(Qt.AlignCenter)
        self.nav_bar.setSpacing(12)
        self.nav_bar.addWidget(self.nav_prev)
        self.nav_bar.addWidget(self.page_label)
        self.nav_bar.addWidget(self.nav_next)

        # Rodapé fixo (fora do ScrollArea)
        self.nav_bar_widget = QWidget()
        self.nav_bar_widget.setAttribute(Qt.WA_TranslucentBackground)
        self.nav_bar_widget.setStyleSheet("background: transparent;")
        self.nav_bar_widget.setLayout(self.nav_bar)

        outer_layout.addWidget(self.nav_bar_widget)

        # garante que o Scroll ocupa o espaço e a nav fica colada embaixo
        outer_layout.setStretch(0, 1)  # scroll_area cresce
        outer_layout.setStretch(1, 0)  # nav fixa

        self.progressBar = QProgressBar()
        self.progressBar.setMaximum(100)
        self.progressBar.setVisible(False)
        self.ajustar_tamanho_tela()
        self.centralizar_janela()

        self.bancos = [
            {"nome": "Asaas", "icone": "imagens/Asaas1.ico",
                "func": self.executar_asaas},
            {"nome": "Banco do Brasil", "icone": "imagens/brasil1.ico",
                "func": self.executar_brasil},
            {"nome": "Nordeste", "icone": "imagens/bnb-icon.ico",
                "func": self.executar_bnb},
            {"nome": "Bradesco", "icone": "imagens/bradesco.ico",
                "func": self.executar_bradesco},
            {"nome": "Btg", "icone": "imagens/btg.ico",
                "func": self.executar_btg},
            {"nome": "Caixa", "icone": "imagens/caixa.ico",
                "func": self.executar_caixa},
            {"nome": "Daycoval", "icone": "imagens/daycoval.ico",
                "func": self.executar_daycoval},
            {"nome": "Inter", "icone": "imagens/icon inter3.ico",
                "func": self.executar_inter},
            {"nome": "Itaú", "icone": "imagens/itau11.ico",
                "func": self.executar_itau},
            {"nome": "Itaú BBA", "icone": "imagens/itau25.ico",
                "func": self.executar_itau2},
            {"nome": "Itaú Manix", "icone": "imagens/itaumanix1.ico",
                "func": self.executar_itaumanix},
            {"nome": "Itaú Consolidado", "icone": "imagens/ItauConsolidado1.ico",
                "func": self.executar_itauconsolidado},
            {"nome": "Nubank", "icone": "imagens/nubank1.ico",
                "func": self.executar_nubank},
            {"nome": "Pagbank", "icone": "imagens/pagbank1.ico",
                "func": self.executar_pagbank},
            {"nome": "Safra", "icone": "imagens/Safra.ico",
                "func": self.executar_safra},
            {"nome": "Santander", "icone": "imagens/santander-br.ico",
                "func": self.executar_santander},
            {"nome": "Sicredi", "icone": "imagens/Sicredi.ico",
                "func": self.executar_sicredi},
            {"nome": "Sofisa", "icone": "imagens/sofisa1.ico",
                "func": self.executar_sofisa},
            {"nome": "Stone", "icone": "imagens/stone.ico",
                "func": self.executar_stone},
            {"nome": "Xp Investimentos", "icone": "imagens/Xp-Investimentos-Logo.ico",
                "func": self.executar_xp}
        ]

        self.bancos.sort(key=lambda x: x["nome"])

        # refs de botões (usadas pelo tema)
        self.button_refs = []

        # 🔹 Paginação
        self.current_page = 0
        self.banks_per_page = 20  # 4 colunas x 5 linhas

        # 🔹 Grid (widget + layout)
        self.grid_widget = QWidget()
        self.grid_widget.setAttribute(Qt.WA_TranslucentBackground)
        self.grid_widget.setSizePolicy(
            QSizePolicy.Preferred, QSizePolicy.MinimumExpanding)

        self.grid_layout = QGridLayout(self.grid_widget)
        self.grid_layout.setContentsMargins(0, 0, 0, 0)
        self.grid_layout.setHorizontalSpacing(32)  # antes 50
        self.grid_layout.setVerticalSpacing(14)    # antes 20

        # Espaçamento proporcional nas colunas/linhas
        for c in range(4):
            self.grid_layout.setColumnStretch(c, 1)
        self.grid_layout.setRowStretch(6, 1)  # espaço extra abaixo

        self.layout.addWidget(self.grid_widget)

        # Renderizar a primeira página
        self.render_page()

    def update_top_offset(self):
        """Atualiza a altura do espaçador superior conforme o tema e o tamanho da janela."""
        if hasattr(self, 'top_offset'):
            fator = self.offset_factor_light if self.light_theme else self.offset_factor_dark
            h = max(0, int(self.height() * fator))
            self.top_offset.changeSize(
                0, h, QSizePolicy.Minimum, QSizePolicy.Fixed)
            # Força o layout a recalcular
            self.layout.invalidate()

    def resizeEvent(self, event):
        if hasattr(self, 'background_label') and self.background_label:
            self.background_label.setGeometry(self.rect())
        self.update_top_offset()  # ⬅️ mantém a grade na posição ideal
        super().resizeEvent(event)

        if hasattr(self, 'top_offset'):
            fator = 0.05
            h = max(0, int(self.height() * fator))
            self.top_offset.changeSize(
                0, h, QSizePolicy.Minimum, QSizePolicy.Fixed)
            self.layout.invalidate()

    def mostrar_confirmacao(self, titulo, mensagem):
        msg = QMessageBox(self)
        msg.setWindowTitle(titulo)
        msg.setText(mensagem)
        msg.setIcon(QMessageBox.Question)
        msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        msg.setDefaultButton(QMessageBox.No)

        # Renomear botões
        btn_sim = msg.button(QMessageBox.Yes)
        btn_nao = msg.button(QMessageBox.No)

        # Aplica estilo visual antes de mostrar a janela
        if self.light_theme:
            msg.setStyleSheet("""
                QMessageBox {
                    background-color: #ffffff;
                }
                QLabel {
                    color: #000000;
                    font-weight: bold;
                    border: none;
                }
            """)
        else:
            msg.setStyleSheet("""
                QMessageBox {
                    background-color: #000000;
                }
                QLabel {
                    color: #ffa500;
                    font-weight: bold;
                    border: none;
                }
            """)

        # Após mostrar, atualiza texto e estilos dos botões
        def estilizar_botoes():
            if btn_sim:
                btn_sim.setText("Sim")
            if btn_nao:
                btn_nao.setText("Não")

            for button in msg.findChildren(QPushButton):
                if self.light_theme:
                    button.setStyleSheet("""
                        QPushButton {
                            background-color: #000000;
                            color: #ffffff;
                            font-weight: bold;
                            padding: 6px 16px;
                            border-radius: 10px;
                            border: 2px solid #ff6600;
                        }
                        QPushButton:hover {
                            background-color: #222222;
                        }
                    """)
                else:
                    button.setStyleSheet("""
                        QPushButton {
                            background-color: #000000;
                            color: #ffa500;
                            font-weight: bold;
                            padding: 6px 16px;
                            border-radius: 10px;
                            border: 2px solid #ffa500;
                        }
                        QPushButton:hover {
                            background-color: #222222;
                        }
                    """)

        QTimer.singleShot(0, estilizar_botoes)

        resultado = msg.exec_()
        return resultado == QMessageBox.Yes

    def toggle_theme(self):
        self.light_theme = not self.light_theme
        if self.light_theme:
            self.apply_light_theme()
        else:
            self.apply_dark_theme()
        self.update_top_offset()

        self.update_top_offset()  # ⬅️ reposiciona a grade para o tema atual

    def add_bank_button(self, label, image_path, function, target_layout=None):
        card = QVBoxLayout()
        card.setSpacing(8)                       # antes 10
        card.setAlignment(Qt.AlignCenter)

        imagem_completa = recurso_caminho(image_path)
        if os.path.exists(imagem_completa):
            pixmap = QPixmap(imagem_completa).scaled(
                28, 28, Qt.KeepAspectRatio, Qt.SmoothTransformation)  # antes 35
        icon_label = QLabel()
        icon_label.setPixmap(pixmap)
        icon_label.setAlignment(Qt.AlignCenter)
        icon_label.setStyleSheet("background: transparent;")
        card.addWidget(icon_label)

        button = QPushButton(label)
        button.setCursor(Qt.PointingHandCursor)
        button.setFixedWidth(130)                # antes 150
        button.setMinimumHeight(34)              # altura mais enxuta
        card.addWidget(button)

        wrapper = QWidget()
        wrapper.setLayout(card)

        button.clicked.connect(lambda: self.executar_banco(function()))
        self.button_refs.append(button)

        if target_layout is not None:
            target_layout.addWidget(wrapper)

        return wrapper

    def executar_banco(self, funcao_banco):
        self.dialog_loader = LoaderDialog(
            self, getattr(self, "light_theme", True))
        self.dialog_loader.show()
        self.loader = LoaderThread()
        self.loader.progress.connect(self.dialog_loader.atualizar_progresso)
        self.loader.finished.connect(
            lambda: self.finalizar_execucao(funcao_banco))
        self.loader.finished.connect(self.dialog_loader.close)
        self.loader.start()

    def finalizar_execucao(self, funcao_banco):
        try:
            funcao_banco(self)
        except Exception as e:
            self.mostrar_mensagem("Erro", str(e))

    def apply_light_theme(self):
        self.setStyleSheet("background-color: #ffffff;")
        self.toggle_btn.setText("🌙 Modo Escuro")
        self.toggle_btn.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                color: #000000;
                padding: 8px 10px;
                font-weight: bold;
                border-radius: 12px;
                border: 2px solid #ff6600;
            }
            QPushButton:hover {
                background-color: rgba(255, 102, 0, 30);
            }
        """)
        self.progressBar.setStyleSheet("""
            QProgressBar {
                border: 2px solid #ff6600;
                border-radius: 10px;
                background-color: #ffffff;
                height: 25px;
                text-align: center;
                font-weight: bold;
                color: #000000;
            }
            QProgressBar::chunk {
                background-color: #ff6600;
                width: 20px;
            }
        """)

        for btn in list(self.button_refs):
            try:
                btn.setStyleSheet("""
                    QPushButton {
                        background-color: #ffffff;  /* melhor legibilidade no escuro */
                        color: #000000;
                        padding: 8px 10px;
                        font-weight: bold;
                        font-size: 12px;
                        border-radius: 12px;
                        border: 2px solid #ffa500;
                }
                QPushButton:hover {
                    background-color: rgba(255,102,0,30);  
                }
            """)
            except RuntimeError:
                self.button_refs.remove(btn)

            # Estilo da navegação (setas/label)
        if hasattr(self, "nav_prev"):
            self.nav_prev.setStyleSheet("""
                QPushButton {
                    background-color: #ffffff;
                    color: #000000;
                    font-weight: bold;
                    border-radius: 10px;
                    border: 2px solid #ffa500;
                }
                QPushButton:disabled {
                    color: #777777;
                    background-color: #cccccc;
                    border: 2px solid #aaaaaa;
                }
                QPushButton:hover {
                    background-color: #dddddd;
                }
            """)
        if hasattr(self, "nav_next"):
            self.nav_next.setStyleSheet("""
                QPushButton {
                    background-color: #ffffff;
                    color: #000000;
                    font-weight: bold;
                    border-radius: 10px;
                    border: 2px solid #ffa500;
                }
                QPushButton:disabled {
                    color: #777777;
                    background-color: #cccccc;
                    border: 2px solid #aaaaaa;
                }
                QPushButton:hover {
                    background-color: #dddddd;
                }
            """)
        if hasattr(self, "page_label"):
            self.page_label.setStyleSheet(
                "color: #000000; font-weight: bold; font-size: 12px;")

        if hasattr(self, "background_label"):
            self.background_label.setPixmap(QPixmap(self.img_light))
            self.background_label.setVisible(True)
            self.background_label.lower()

    def apply_dark_theme(self):
        self.setStyleSheet("background-color: #000000;")
        self.titulo.setStyleSheet("color: #ffa500;")
        self.toggle_btn.setText("🌞 Modo Claro")
        self.toggle_btn.setStyleSheet("""
            QPushButton {
                background-color: #ffffff;
                color: #000000;
                padding: 8px 10px;
                font-weight: bold;
                border-radius: 12px;
                border: 2px solid #ffa500;
            }
            QPushButton:hover {
                background-color: #dddddd;
            }
        """)
        self.progressBar.setStyleSheet("""
            QProgressBar {
                border: 2px solid #ffa500;
                border-radius: 10px;
                background-color: #222222;
                height: 25px;
                text-align: center;
                font-weight: bold;
                color: #ffffff;
            }
            QProgressBar::chunk {
                background-color: #ffa500;
                width: 20px;
            }
        """)

        for btn in list(self.button_refs):
            try:
                btn.setStyleSheet("""
                    QPushButton {
                        background-color: #000000;
                        color: #ffffff;
                        padding: 8px 10px;
                        font-weight: bold;
                        font-size: 12px;
                        border-radius: 12px;
                        border: 2px solid #ffffff;
                    }
                    QPushButton:hover {
                        background-color: #222222;
                    }
                """)
            except RuntimeError:
                # Botão já foi destruído – remove da lista para não tentar novamente
                self.button_refs.remove(btn)

            # Estilo da navegação (setas/label)
        if hasattr(self, "nav_prev"):
            self.nav_prev.setStyleSheet("""
                QPushButton {
                    background-color: #ffffff;
                    color: #000000;
                    font-weight: bold;
                    border-radius: 12px;
                    border: 2px solid #ffa500;
                }
                QPushButton:disabled {
                    color: #777777;
                    background-color: #cccccc;
                    border: 2px solid #aaaaaa;
                }
                QPushButton:hover {
                    background-color: #dddddd;
                }
            """)
        if hasattr(self, "nav_next"):
            self.nav_next.setStyleSheet("""
                QPushButton {
                    background-color: #ffffff;
                    color: #000000;
                    font-weight: bold;
                    border-radius: 12px;
                    border: 2px solid #ffa500;
                }
                QPushButton:disabled {
                    color: #777777;
                    background-color: #cccccc;
                    border: 2px solid #aaaaaa;
                }
                QPushButton:hover {
                    background-color: #dddddd;
                }
            """)
        if hasattr(self, "page_label"):
            self.page_label.setStyleSheet("color: #ffa500; font-weight: bold;")

        if hasattr(self, "background_label"):
            self.background_label.setPixmap(QPixmap(self.img_dark))
            # agora também mostra no escuro
            self.background_label.setVisible(True)
            self.background_label.lower()

    def mostrar_mensagem(self, titulo, mensagem):
        dialog = QDialog(self)
        dialog.setWindowTitle(titulo)
        dialog.setFixedSize(500, 180)
        dialog.setWindowFlags(dialog.windowFlags() | Qt.WindowCloseButtonHint)

        layout = QVBoxLayout(dialog)

        text_area = QTextEdit()
        text_area.setText(mensagem)
        text_area.setReadOnly(True)
        layout.addWidget(text_area)

        botoes = QDialogButtonBox(QDialogButtonBox.Ok)
        botoes.accepted.connect(dialog.accept)
        layout.addWidget(botoes)

        if self.light_theme:
            dialog.setStyleSheet("""
                QDialog {
                    background-color: #ffffff;
                }
                QTextEdit {
                    background-color: #ffffff;
                    color: #000000;
                    font-weight: bold;
                    border: none;
                }
                QDialogButtonBox QPushButton {
                    background-color: #000000;
                    color: white;
                    padding: 6px 16px;
                    border-radius: 8px;
                    font-weight: bold;
                }
                QDialogButtonBox QPushButton:hover {
                    background-color: #333333;
                }
            """)
        else:
            dialog.setStyleSheet("""
                QDialog {
                    background-color: #000000;
                    color: #ffa500;
                }
                QTextEdit {
                    background-color: #000000;
                    color: #ffa500;
                    font-weight: bold;
                    border: none;
                }
                QDialogButtonBox QPushButton {
                    background-color: #ffffff;
                    color: #ffa500;
                    padding: 6px 16px;
                    border-radius: 8px;
                    font-weight: bold;
                }
                QDialogButtonBox QPushButton:hover {
                    background-color: #dddddd;
                }
            """)

        dialog.exec_()

    def executar_asaas(self):
        from Asaas import processar_pdf_custom
        return processar_pdf_custom

    def executar_brasil(self):
        from Brasil import processar_pdf_custom
        return processar_pdf_custom

    def executar_bnb(self):
        from BNB import processar_pdf_custom
        return processar_pdf_custom

    def executar_bradesco(self):
        from Bradesco import processar_pdf_custom
        return processar_pdf_custom

    def executar_btg(self):
        from Btg import processar_pdf_custom
        return processar_pdf_custom

    def executar_caixa(self):
        from Caixa import processar_pdf_custom
        return processar_pdf_custom

    def executar_daycoval(self):
        from Daycoval import processar_pdf_custom
        return processar_pdf_custom

    def executar_inter(self):
        from Inter import processar_pdf_custom
        return processar_pdf_custom

    def executar_pagbank(self):
        from Pagbank import processar_pdf_custom
        return processar_pdf_custom

    def executar_itau(self):
        from itau import processar_pdf_custom
        return processar_pdf_custom

    def executar_itau2(self):
        from Itau2 import processar_pdf_custom
        return processar_pdf_custom

    def executar_itaumanix(self):
        from itau_MANIX import processar_pdf_custom
        return processar_pdf_custom

    def executar_itauconsolidado(self):
        from ItauConsolidado import processar_pdf_custom
        return processar_pdf_custom

    def executar_nubank(self):
        from Nubank import processar_pdf_custom
        return processar_pdf_custom

    def executar_safra(self):
        from Safra import processar_pdf_custom
        return processar_pdf_custom

    def executar_santander(self):
        from Santander import processar_pdf_custom
        return processar_pdf_custom

    def executar_sicredi(self):
        from Sicredi import processar_pdf_custom
        return processar_pdf_custom

    def executar_sofisa(self):
        from Sofisa import processar_pdf_custom
        return processar_pdf_custom

    def executar_stone(self):
        from Stone import processar_pdf_custom
        return processar_pdf_custom

    def executar_xp(self):
        from XpInvestimentos import processar_pdf_custom
        return processar_pdf_custom

    def ajustar_tamanho_tela(self):
        screen = QDesktopWidget().screenGeometry()
        largura = min(983, screen.width() - 100)
        altura = min(750, screen.height() - 100)
        self.resize(largura, altura)

    def centralizar_janela(self):
        frame_geometry = self.frameGeometry()
        centro_tela = QDesktopWidget().availableGeometry().center()
        frame_geometry.moveCenter(centro_tela)
        self.move(frame_geometry.topLeft())

    def clear_grid(self):
        while self.grid_layout.count():
            item = self.grid_layout.takeAt(0)
            w = item.widget()
            if w is not None:
                w.deleteLater()

    def total_pages(self):
        return max(1, (len(self.bancos) + self.banks_per_page - 1) // self.banks_per_page)

    def clear_layout(self, layout):
        while layout.count():
            item = layout.takeAt(0)
            w = item.widget()
            child_layout = item.layout()
            if w is not None:
                w.deleteLater()
            elif child_layout is not None:
                self.clear_layout(child_layout)

    def render_page(self):
        # Zera refs e limpa o grid
        self.button_refs = []
        self.clear_grid()

        start = self.current_page * self.banks_per_page
        end = min(start + self.banks_per_page, len(self.bancos))
        bancos_pagina = self.bancos[start:end]

        cols = 4
        row = 0
        col = 0

        for banco in bancos_pagina:
            w = self.add_bank_button(
                banco["nome"], banco["icone"], banco["func"], target_layout=None)
            self.grid_layout.addWidget(w, row, col, alignment=Qt.AlignCenter)
            col += 1
            if col >= cols:
                col = 0
                row += 1

        self.update_nav()

        # Reaplicar tema aos novos botões
        if getattr(self, "light_theme", True):
            self.apply_light_theme()
        else:
            self.apply_dark_theme()

    def update_nav(self):
        total = self.total_pages()
        self.page_label.setText(f"Página {self.current_page + 1} / {total}")
        self.nav_prev.setEnabled(self.current_page > 0)
        self.nav_next.setEnabled(self.current_page < total - 1)

    def next_page(self):
        if self.current_page < self.total_pages() - 1:
            self.current_page += 1
            self.render_page()

    def prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.render_page()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    janela = BancoApp()
    janela.show()
    sys.exit(app.exec_())
