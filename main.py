# Importação de bibliotecas necessárias
import requests
import os
import base64
import json
import argparse
import configparser
import sys
import subprocess
from PyQt5.QtCore import Qt, QTimer, QThread, pyqtSignal, QEventLoop
from PyQt5.QtGui import QMovie, QPixmap, QIcon
from PyQt5.QtWidgets import QApplication, QLabel, QVBoxLayout, QWidget, QMessageBox
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders
import io
from datetime import datetime, timedelta
from pytz import utc
import xlsxwriter
import shutil
from PyQt5.QtWidgets import (
    QMainWindow, QPushButton, QDialog,
    QCalendarWidget, QTimeEdit, QLineEdit, QFileDialog, QTextEdit,
    QDesktopWidget, QSpacerItem, QSizePolicy, QSpinBox, QButtonGroup, QRadioButton
)
from PyQt5.QtGui import QFont

# Definição de variáveis para armazenamento de caminhos de arquivos e recursos
arquivo_dados = os.path.join(os.path.dirname(__file__),"resources/dados_api.json")
arquivo_config = os.path.join(os.path.dirname(__file__),"resources/config.ini")
urmobo = os.path.join(os.path.dirname(__file__), "resources/urmobo.png")

# Função para obter dados da API, com lógica para leitura de cache local
def obter_dados_da_api(usuario, senha):
    # Tentativa de leitura dos dados da API do cache local
    try:
        with open(arquivo_dados, 'r') as arquivo:
            dados = json.load(arquivo)
            return dados
    except FileNotFoundError:
        # Caso o cache não exista, realizar a solicitação à API
        url = 'https://integracao.urmobo.com.br/equipamentos'

        # Criação do cabeçalho de autenticação
        credenciais = f"{usuario}:{senha}"
        credenciais_base64 = base64.b64encode(credenciais.encode()).decode()
        cabecalho_autenticacao = {'Authorization': f'Basic {credenciais_base64}'}

        # Realização da solicitação GET à API
        response = requests.get(url, headers=cabecalho_autenticacao)

        # Tratamento da resposta
        if response.status_code == 200:
            dados = response.json()
            with open(arquivo_dados, 'w') as arquivo:
                json.dump(dados, arquivo)
            return dados
        else:
            raise Exception(f"Erro na solicitação: {response.status_code} - {response.text}")

# Classe para a janela de introdução do aplicativo
class Introducao(QDialog):
    def __init__(self):
        super().__init__()

        # Configurações iniciais da janela
        self.setWindowTitle("Introdução")
        self.setGeometry(100, 100, 390, 340)
        self.setWindowFlags(Qt.FramelessWindowHint)  # Remove a borda da janela
        self.setAttribute(Qt.WA_TranslucentBackground)

        # Configuração do ícone da janela
        icon_pixmap = QPixmap(os.path.join(os.path.dirname(__file__), "resources/titulo.png"))
        self.setWindowIcon(QIcon(icon_pixmap))

        # Centralizar a janela na tela
        self.centralizar_na_tela()

        # Layout da janela
        layout = QVBoxLayout(self)

        # Configuração da exibição do gif
        self.label = QLabel(self)
        self.label.setAlignment(Qt.AlignCenter)
        self.set_gif(os.path.join(os.path.dirname(__file__),"resources/video.gif"))
        layout.addWidget(self.label)

        # Temporizador para fechar a janela após um intervalo
        QTimer.singleShot(5000, self.abrir_janela_principal)

    def centralizar_na_tela(self):
        # Lógica para centralizar a janela na tela
        screen_geometry = QDesktopWidget().screenGeometry()
        x = (screen_geometry.width() - self.width()) // 2
        y = (screen_geometry.height() - self.height()) // 2
        self.move(x, y)

    def set_gif(self, gif_path):
        # Configuração do gif na QLabel
        movie = QMovie(gif_path)
        self.label.setMovie(movie)
        movie.start()

    def abrir_janela_principal(self):
        # Método para fechar a janela de introdução
        self.close()

# Classe para a janela de configurações do aplicativo
class ConfiguracoesDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)

        # Configurações iniciais da janela de diálogo
        self.setFixedSize(220, 250)
        self.setWindowTitle("Configurações")
        self.setModal(True)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        # Centraliza a janela de diálogo
        parent_geometry = parent.geometry() if parent else QApplication.desktop().screenGeometry()
        self.setGeometry(
            parent_geometry.center().x() - self.width() // 2,
            parent_geometry.center().y() - self.height() // 2,
            self.width(),
            self.height()
        )

        # Layout para a janela de diálogo
        layout = QVBoxLayout(self)

        # Opções de configuração disponíveis
        opcoes = ["DIAS", "DADOS API", "DADOS DO USUÁRIO", "DESK", "TERMOS", "CONTATOS", "SOBRE"]
        for opcao in opcoes:
            button = self.create_button(opcao, self.abrir_janela_opcao)
            layout.addWidget(button)

    def create_button(self, text, on_click):
        # Cria um botão estilizado para a janela de diálogo
        button = QPushButton(text, self)
        button.setStyleSheet("text-align: center; color: #143E79; font-family: 'Monomaniac One'; font-size: 9pt;")
        button.setCursor(Qt.PointingHandCursor)
        button.clicked.connect(on_click)
        return button

    def abrir_janela_opcao(self):
        # Função chamada quando um botão de opção é clicado
        sender_button = self.sender()
        opcao_dialog = OpcaoDialog(self, sender_button.text())
        opcao_dialog.exec_()

# Classe para a janela de opção do aplicativo
class OpcaoDialog(QDialog):
    def __init__(self, parent=None, opcao=None):
        super().__init__(parent)

        # Configurações iniciais da janela de opção
        self.setFixedSize(220, 250)
        self.setWindowTitle(f"{opcao}")
        self.setModal(True)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        # Layout para a janela de opção
        layout = QVBoxLayout(self)

        # Inicializa widgets de usuário, senha e dias
        self.usuario_edit = None
        self.senha_edit = None
        self.dias_spinbox = None

        # Estrutura condicional para adicionar widgets específicos de cada opção
        if opcao == "DIAS":
           self.add_timer_widgets(layout)
           self.carregar_configuracoes_timer()
        elif opcao in ["DADOS API", "DADOS DO USUÁRIO"]:
            self.add_data_entry_widgets(layout, opcao)
        elif opcao in ["TERMOS", "CONTATOS"]:
            self.add_file_selection_widgets(layout)
        elif opcao == "DESK":
            self.add_desk_widgets(layout)
        elif opcao == "SOBRE":
            self.add_about_text_widget(layout)

    def add_timer_widgets(self, layout):
        calendar_widget = QCalendarWidget(self)
        calendar_widget.setMinimumSize(200, 150)
        time_edit = QTimeEdit(self)
        layout.addWidget(calendar_widget)
        layout.addWidget(time_edit)

    def add_data_entry_widgets(self, layout, opcao):
        usuario_label_text = "Usuário:"
        senha_label_text = "Senha:"

        if opcao == "DADOS API":
            usuario_label_text = "Usuário API:"
            senha_label_text = "Senha API:"

        usuario_label = QLabel(usuario_label_text, self)
        self.usuario_edit = QLineEdit(self)
        self.usuario_edit.setObjectName("usuario_edit")  # Adicione o objectName aqui
        senha_label = QLabel(senha_label_text, self)
        self.senha_edit = QLineEdit(self)
        self.senha_edit.setObjectName("senha_edit")  # Adicione o objectName aqui
        self.senha_edit.setEchoMode(QLineEdit.Password)
        salvar_button = self.create_styled_button("Salvar", callback=self.salvar_dados_api if opcao == "DADOS API" else self.salvar_dados)
        self.add_input_widgets(layout, [usuario_label, self.usuario_edit, senha_label, self.senha_edit, salvar_button])

        # Tentar carregar as informações salvas do arquivo config.ini
        try:
            config = configparser.ConfigParser()
            config.read(os.path.join(os.path.dirname(__file__),'resources/config.ini'))

            if opcao == "DADOS API":
                if 'API' in config and 'Usuario' in config['API']:
                    usuario_salvo_api = config['API']['Usuario']
                    self.usuario_edit.setText(usuario_salvo_api)

                if 'API' in config and 'Senha' in config['API']:
                    senha_salva_api = config['API']['Senha']
                    self.senha_edit.setText(senha_salva_api)

            elif opcao == "DADOS DO USUÁRIO":
                if 'USER' in config and 'Usuario2' in config['USER']:
                    usuario_salvo_user = config['USER']['Usuario2']
                    self.usuario_edit.setText(usuario_salvo_user)

                if 'USER' in config and 'Senha2' in config['USER']:
                    senha_salva_user = config['USER']['Senha2']
                    self.senha_edit.setText(senha_salva_user)

        except Exception as e:
            print(f"Erro ao carregar informações salvas: {e}")

    def add_file_selection_widgets(self, layout):
        arquivo_label = QLabel("Selecione o arquivo:", self)
        arquivo_button = self.create_styled_button("Escolher Arquivo", callback=self.selecionar_arquivo)
        self.add_input_widgets(layout, [arquivo_label, arquivo_button])

    def add_desk_widgets(self, layout):
        # Label para a mensagem
        texto = (
            "Deseja realizar o envio do relatório via abertura de chamado?"
        )

        texto_edit = QTextEdit(self)
        texto_edit.setPlainText(texto)
        texto_edit.setStyleSheet(
            "color: #143E79; font-family: 'Monomaniac One'; font-size: 10pt; background-color: #ffffff;"
        )
        texto_edit.setAlignment(Qt.AlignCenter)
        texto_edit.setReadOnly(True)

        layout.addWidget(texto_edit)

        # Grupo de RadioButtons
        self.desk_radio_group = QButtonGroup(self)

        # RadioButton para a opção "Sim"
        self.sim_radio = QRadioButton("Sim", self)
        self.desk_radio_group.addButton(self.sim_radio)

        # RadioButton para a opção "Não"
        self.nao_radio = QRadioButton("Não", self)
        self.desk_radio_group.addButton(self.nao_radio)

        layout.addWidget(self.sim_radio)
        layout.addWidget(self.nao_radio)

        self.sim_radio.toggled.connect(self.salvar_configuracao_desk)
        self.nao_radio.toggled.connect(self.salvar_configuracao_desk)

        self.carregar_configuracao_desk()

        self.carregar_configuracao_desk()

    def salvar_configuracao_desk(self):
        escolha = "Sim" if self.sim_radio.isChecked() else "Nao" if self.nao_radio.isChecked() else ""
        config = configparser.ConfigParser()
        config.read(arquivo_config)
        if 'DESK' not in config:
            config.add_section('DESK')
        config['DESK']['escolha'] = escolha
        with open(arquivo_config, 'w') as configfile:
            config.write(configfile)

    def carregar_configuracao_desk(self):
        config = configparser.ConfigParser()
        config.read(arquivo_config)
        if 'DESK' in config and 'escolha' in config['DESK']:
            escolha = config['DESK']['escolha']
            if escolha == "Sim":
                self.sim_radio.setChecked(True)
            elif escolha == "Nao":
                self.nao_radio.setChecked(True)

    def add_about_text_widget(self, layout):
        sobre_texto = (
            "O LogMobo foi desenvolvido com o objetivo de facilitar o registro de aparelhos inativos no grupo SIRTEC."
            " Em caso de eventuais erros ou necessidade de configurações adicionais, não hesite em entrar em contato"
            " com Vinícius Kirinus para obter assistência personalizada."
        )

        sobre_texto_edit = QTextEdit(self)
        sobre_texto_edit.setPlainText(sobre_texto)
        sobre_texto_edit.setStyleSheet(
            "color: #143E79; font-family: 'Monomaniac One'; font-size: 10pt; background-color: #ffffff;"
        )
        sobre_texto_edit.setAlignment(Qt.AlignCenter)
        sobre_texto_edit.setReadOnly(True)

        layout.addWidget(sobre_texto_edit)

    def add_input_widgets(self, layout, widgets):
        quebra_de_linha = QLabel("", self)
        layout.addWidget(quebra_de_linha)

        for widget in widgets:
            layout.addWidget(widget)

        spacer_item = QSpacerItem(40, 50, QSizePolicy.Minimum, QSizePolicy.Expanding)
        layout.addItem(spacer_item)

    def create_styled_button(self, text, callback=None):
        button = QPushButton(text, self)
        button.setStyleSheet("background-color: #85CAF8; color: #143E79; "
                             "border: 2px solid #85CAF8; border-radius: 10px; font-family: 'Monomaniac One';")
        if callback:
            button.clicked.connect(callback)
        return button
    
    def add_timer_widgets(self, layout):
        # Adiciona um QTextEdit para a mensagem diretamente ao layout
        mensagem_textedit = QTextEdit("Selecione a quantidade de dias inativo para busca:", self)
        mensagem_textedit.setReadOnly(True)
        mensagem_textedit.setAlignment(Qt.AlignCenter)
        mensagem_textedit.setStyleSheet("font-family: 'Monomaniac One'; font-size: 10pt; border: 0px; color: #143E79")

        # Adiciona o QSpinBox para selecionar a quantidade de dias
        self.dias_spinbox = QSpinBox(self)
        self.dias_spinbox.setObjectName("dias_spinbox")
        self.dias_spinbox.setMinimum(1)
        self.dias_spinbox.setMaximum(30)

        salvar_button = self.create_styled_button("Salvar", callback=self.salvar_timer)
        self.add_input_widgets(layout, [mensagem_textedit, self.dias_spinbox, salvar_button])
        
    def carregar_configuracoes_timer(self):
        try:
            config = configparser.ConfigParser()
            config.read(arquivo_config)

            if 'DATE' in config and 'data' in config['DATE']:
                quantidade_dias = int(config['DATE']['data'])
                self.dias_spinbox.setValue(quantidade_dias)

        except Exception as e:
            print(f"Erro ao carregar configurações do timer: {e}")

    def salvar_timer(self):
        if self.dias_spinbox is not None:
            quantidade_dias = self.dias_spinbox.value()

            try:
                config = configparser.ConfigParser()
                config.read(os.path.join(os.path.dirname(__file__),'config.ini'))

                if 'DATE' not in config:
                    config.add_section('DATE')

                config['DATE']['data'] = str(quantidade_dias)

                with open(arquivo_config, 'w') as configfile:
                    config.write(configfile)

                QMessageBox.information(self, "Sucesso", "Quantidade de dias salva!")
            except Exception as e:
                print(f"Erro ao salvar quantidade de dias: {e}")
        else:
            QMessageBox.warning(self, "Aviso", "Não foi possível encontrar o campo para a quantidade de dias.")

    def salvar_dados_api(self):
        if self.usuario_edit is not None and self.senha_edit is not None:
            usuario = self.usuario_edit.text()
            senha = self.senha_edit.text()

            try:
                config = configparser.ConfigParser()
                config.read(arquivo_config)

                if 'API' not in config:
                    config.add_section('API')

                config['API']['usuario'] = usuario
                config['API']['senha'] = senha

                with open(arquivo_config, 'w') as configfile:
                    config.write(configfile)

                QMessageBox.information(self, "Sucesso", "Dados API salvos!")
            except Exception as e:
                print(f"Erro ao salvar dados API: {e}")
        else:
            QMessageBox.warning(self, "Aviso", "Não foi possível encontrar os campos de usuário e senha.")

    def salvar_dados(self):
        if self.usuario_edit is not None and self.senha_edit is not None:
            usuario2 = self.usuario_edit.text()
            senha2 = self.senha_edit.text()

            try:
                config = configparser.ConfigParser()
                config.read(arquivo_config)

                if 'USER' not in config:
                    config.add_section('USER')

                config['USER']['usuario2'] = usuario2
                config['USER']['senha2'] = senha2

                with open(arquivo_config, 'w') as configfile:
                    config.write(configfile)

                QMessageBox.information(self, "Sucesso", "Dados do usuário salvos!")
            except Exception as e:
                print(f"Erro ao salvar dados API: {e}")
        else:
            QMessageBox.warning(self, "Aviso", "Não foi possível encontrar os campos de usuário e senha.")

    def selecionar_arquivo(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_dialog = QFileDialog(self)
        file_dialog.setOptions(options)
        file_dialog.setNameFilter("Arquivos (*.xlsx);;Todos os arquivos (*)")
        file_dialog.setWindowTitle("Selecione o arquivo")
        file_dialog.setFileMode(QFileDialog.ExistingFile)

        if file_dialog.exec_() == QFileDialog.Accepted:
            selected_file = file_dialog.selectedFiles()[0]
            print(f"Arquivo selecionado: {selected_file}")
            self.salvar_arquivo(selected_file)

    def salvar_arquivo(self, selected_file):
        file_name = os.path.basename(selected_file)
        current_directory = os.path.dirname(os.path.abspath(__file__))
        current_file_path = os.path.join(current_directory, file_name)

        try:
            shutil.copy(selected_file, current_file_path)
            QMessageBox.information(self, "Sucesso", "Arquivo salvo com sucesso!")
        except Exception as e:
            QMessageBox.critical(self, "Erro", "Erro ao salvar o arquivo!")

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setGeometry(100, 100, 390, 340)
        self.setFixedSize(390, 340)

        self.envio_thread = None

        self.mostrar_introducao()

        background_pixmap = QPixmap(os.path.join(os.path.dirname(__file__), "resources/fundo.png"))
        background_label = QLabel(self)
        background_label.setPixmap(background_pixmap)
        background_label.setGeometry(0, 0, 390, 325)
        icon_pixmap = QPixmap(os.path.join(os.path.dirname(__file__), "resources/titulo.png"))
        self.setWindowIcon(QIcon(icon_pixmap))

        button_iniciar = QPushButton("Iniciar", self)
        button_iniciar.setGeometry(125, 230, 140, 40)
        button_iniciar.setStyleSheet(
            "background-color: #348CC7; color: white; border: 2px solid #348CC7; border-radius: 20px;"
        )
        button_iniciar.setFont(QFont("Monomaniac One", 12))
        button_iniciar.clicked.connect(self.iniciar_processo)
        button_iniciar.setCursor(Qt.PointingHandCursor)

        button_configurar = QPushButton("Configurações", self)
        button_configurar.setGeometry(146, 280, 95, 30)
        button_configurar.setStyleSheet(
            "background-color: #348CC7; color: white; border: 2px solid #348CC7; border-radius: 15px;"
        )
        button_configurar.setFont(QFont("Monomaniac One", 9))
        button_configurar.clicked.connect(self.abrir_configuracoes)
        button_configurar.setCursor(Qt.PointingHandCursor)

        self.setWindowTitle("LOGMOBO")

    def iniciar_processo(self):
        self.envio_thread = Loading()
        self.envio_thread.show()
        self.envio_thread.centralizar_na_tela()

    def mostrar_introducao(self):
        intro_dialog = Introducao()
        intro_dialog.show()
        self.centralizar_na_tela()
        
    def centralizar_na_tela(self):
        screen_geometry = QDesktopWidget().screenGeometry()
        x = (screen_geometry.width() - self.width()) // 2
        y = (screen_geometry.height() - self.height()) // 2
        self.move(x, y)

    def abrir_configuracoes(self):
        configuracoes_dialog = ConfiguracoesDialog(self)
        configuracoes_dialog.exec_()

class EnvioThread(QThread):
    enviado_com_sucesso = pyqtSignal()
    falha_no_envio = pyqtSignal(str)

    def run(self):
        def obter_credenciais_do_arquivo_config():
            config = configparser.ConfigParser()
            config.read(arquivo_config)

            if 'API' in config:
                usuario = config['API'].get('usuario')
                senha = config['API'].get('senha')

                if usuario and senha:
                    return usuario, senha

            raise Exception("Credenciais não encontradas no arquivo de configuração.")
        try:
            # ENVIO DOS E-MAILS
            # Defina arquivo_config globalmente
            arquivo_config = os.path.join(os.path.dirname(__file__),"resources/config.ini")

            # Função para registrar mensagens no arquivo de log
            def registrar_log(mensagem):
                data_hora = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                registro = f'{data_hora} - {mensagem}\n'
                with open(arquivo_log, 'a') as log:
                    log.write(registro)


            # Obter os dados da API
            usuario, senha = obter_credenciais_do_arquivo_config()
            dados_da_api = obter_dados_da_api(usuario, senha)

            def obter_credenciais_do_arquivo_config():
                config = configparser.ConfigParser()
                config.read(arquivo_config)

                if 'USER' in config:
                    username = config['USER'].get('usuario2')
                    senha = config['USER'].get('senha2')
                    
                    # Remover aspas duplas da senha, se presentes
                    username = username.strip('"')
                    senha = senha.strip('"')

                    print(f'Username lido do config.ini: {username}')
                    print(f'Senha lida do config.ini: {senha}')

                if 'DESK' in config:
                    escolha = config['DESK'].get('escolha')
                    print(f'Escolha lida do config.ini: {escolha}')

                return username, senha, escolha
            
            # Configurações de e-mail 
            servidor_email = "" #necessário informar servidor de e-mail
            porta = 465 # Porta para conexão SSL (alterar caso necessário)
            username, senha, escolha = obter_credenciais_do_arquivo_config()


            # Inicialize a lista de registro_aparelhos_inativos
            registro_aparelhos_inativos = []


            # Criar um DataFrame a partir dos dados da API
            df = pd.DataFrame(dados_da_api)

            # Selecionar apenas as colunas desejadas
            colunas_desejadas = ["Id", "UltimaComunicacao", "Entidade", "isAtivo"]
            df = df[colunas_desejadas]

            # Aplicar o filtro na coluna "Entidade"
            df["Entidade"] = df["Entidade"].str.replace("Produção ", "").str.replace("Apoio ", "")

            # Nome do arquivo de log
            arquivo_log = 'log.txt'

            # Pré-processamento: converter a coluna 'UltimaComunicacao' em datetime com informações de fuso horário (UTC)
            df['UltimaComunicacao'] = pd.to_datetime(df['UltimaComunicacao'], errors='coerce')

            # Remover linhas onde a data não pôde ser formatada corretamente
            df = df.dropna(subset=['UltimaComunicacao'])

            # Ler o valor de dias_atras do arquivo config.ini
            def obter_quantidade_dias_atras():
                config = configparser.ConfigParser()
                config.read(os.path.join(os.path.dirname(__file__),"resources/config.ini"))

                if 'DATE' in config:
                    dias_atras = config['DATE'].getint('data', fallback=7)
                    print(f'Quantidade de dias a subtrair lida do config.ini: {dias_atras}')
                    return dias_atras
                else:
                    # Se a seção DATE não estiver presente, use 7 como valor padrão
                    print('Seção [DATE] não encontrada no config.ini. Usando 7 como valor padrão.')
                    return 7

            # Calcular data há uma quantidade de dias atrás lida do config.ini
            quantidade_dias_atras = obter_quantidade_dias_atras()
            data_semana_atras = datetime.now() - timedelta(days=quantidade_dias_atras)

            # Filtrar aparelhos com mais de uma semana sem comunicação e com isAtivo igual a True
            aparelhos_inativos = df[(df['UltimaComunicacao'] <= pd.to_datetime(data_semana_atras, utc=True)) & (df['isAtivo'] == True)]

            # Leitura das planilhas TERMOS.xlsx e CONTATOS.xlsx -> NECESSÁRIO INCLUIR NAS CONFIGURAÇÕES (TERMOS/CONTATOS DO FIGMA)
            termos = pd.read_excel(os.path.join(os.path.dirname(__file__), "resources/TERMOS.xlsx"))        # Substitua 'Termos.xlsx' pelo nome do seu arquivo de termos
            contatos = pd.read_excel(os.path.join(os.path.dirname(__file__), "resources/CONTATOS.xlsx"))    # Substitua 'Contatos.xlsx' pelo nome do seu arquivo de contatos


            # Iterar pelos aparelhos inativos
            for index, aparelho in aparelhos_inativos.iterrows():
                id_aparelho = aparelho['Id']
                unidade = aparelho['Entidade']

                print(f'Processando aparelho ID: {id_aparelho}, Unidade: {unidade}')

                # Encontrar o colaborador responsável
                termo = termos[termos['Equipamento'] == id_aparelho]
                if not termo.empty:
                    colaborador = termo['Nome'].values[0]
                else:
                    colaborador = "Não encontrado"

                # Encontrar o e-mail do colaborador responsável na coluna "Email"
                contato = contatos[contatos['Unidade'] == unidade]
                if not contato.empty:
                    emails_colaboradores = contato['Email'].str.split(', ').tolist()
                else:
                    emails_colaboradores = []

                if emails_colaboradores:
                    for email_colaborador in emails_colaboradores:
                        # Configurar e enviar o e-mail
                        de = 'ti@sirtec.com.br'
                        para = email_colaborador
                        assunto = f"REGISTRO DE APARELHO INATIVO - {id_aparelho}"
                        print(assunto)
                        corpo = f"""
                            <!DOCTYPE html>
                            <html>
                            <head>
                            <style>
                            body {{
                                text-align: center;
                                margin: 0;
                                padding: 0;
                            }}
                            h3 {{
                                font-weight: bold;
                            }}
                            .header {{
                                text-align: center;
                            }}
                            .header img {{
                                width: 100px;
                                height: 100px;
                                display: block;
                                margin: 0 auto;
                            }}
                            .content {{
                                text-align: center;
                            }}
                            .footer {{
                                font-style: italic;
                            }}
                            </style>
                            </head>
                            <body>
                            <div class="header">
                                <h3>APARELHO INATIVO IDENTIFICADO NO SISTEMA</h3>
                            </div>
                            <div class="content">
                                <p>SEGUE ABAIXO OS DADOS PARA VERIFICAÇÃO:</p>
                                <p><strong>ID:</strong> {id_aparelho}</p>
                                <p><strong>RESPONSAVEL:</strong> {colaborador}</p>
                            </div>
                            <div class="footer">
                                Mensagem de envio automático, por favor não responda a mesma.
                            </div>
                            </body>
                            </html>
                            """.replace('\n', '')  # Remova todas as ocorrências de \n

                        mensagem = MIMEMultipart()
                        mensagem['De'] = de
                        mensagem['Para'] = ', '.join(para)
                        mensagem['Subject'] = assunto
                        
                        # Anexar a imagem ao e-mail com o Content-ID correspondente
                        with open(urmobo, 'rb') as imagem:
                            imagem_anexada = MIMEImage(imagem.read())
                            imagem_anexada.add_header('Content-ID', '<urmobo.png>')
                            mensagem.attach(imagem_anexada)

                        mensagem.attach(MIMEText(corpo, 'html'))

                        try:
                            server = smtplib.SMTP_SSL(servidor_email, porta)
                            server.login(username, senha)
                            texto_do_email = mensagem.as_string()
                            server.sendmail(de, para, texto_do_email)
                            print(f'E-mail enviado com sucesso para {para}')

                            # Registro de sucesso no log
                            registro_log = f'E-mail enviado com sucesso para {para}. ID: {id_aparelho}, Unidade: {unidade}'
                            registrar_log(registro_log)

                            # Adicione o registro do aparelho inativo à lista
                            registro_aparelhos_inativos.append(f"ID: {id_aparelho}, Unidade: {unidade}, Responsável: {colaborador}")

                        except Exception as e:
                            print(f'Erro ao enviar o e-mail para {para}: {str(e)}')

                            # Registro de falha no log
                            registro_log = f'Erro ao enviar o e-mail para {para}: {str(e)}. ID: {id_aparelho}, Unidade: {unidade}'
                            registrar_log(registro_log)

            # Comando de impressão para indicar o término do envio de e-mails
            print("Envio de e-mails concluído.")
            registrar_log("Envio de e-mails concluído.")

            if escolha == "Sim":
                # Criar um DataFrame a partir da lista registro_aparelhos_inativos
                df_relatorio = pd.DataFrame({'Aparelhos Inativos': registro_aparelhos_inativos})

                # Nome do arquivo Excel de relatório
                arquivo_relatorio = os.path.join(os.path.dirname(__file__), 'aparelhos_inativos.xlsx')

                # Salvar o DataFrame em um arquivo Excel
                df_relatorio.to_excel(arquivo_relatorio, sheet_name='aparelhos_inativos', index=False)

                # Envie o registro de aparelhos inativos por e-mail
            
                try:
                    server = smtplib.SMTP_SSL(servidor_email, porta)
                    server.login(username, senha)

                    de = ''  # Substitua pelo seu e-mail
                    para = ''  # Substitua pelo e-mail de destino
                    assunto = "LOG DE APARELHOS INATIVOS"

                    corpo_relatorio = f"""
                    <!DOCTYPE html>
                    <html>
                    <head>
                    <style>
                    body {{
                        text-align: center;
                        margin: 0;
                        padding: 0;
                    }}
                    h3 {{
                        font-weight: bold;
                    }}
                    .header {{
                        text-align: center;
                    }}
                    .content {{
                        text-align: center;
                    }}
                    .footer {{
                        font-style: italic;
                    }}
                    </style>
                    </head>
                    <body>
                    <div class="header">
                        <h3>REGISTRO DE APARELHOS INATIVOS</h3>
                    </div>
                    <div class="content">
                        <p>Abaixo está o registro de aparelhos identificados como inativos:</p>
                        <ul>
                            {"<li>" + "</li><li>".join(registro_aparelhos_inativos) + "</li>"}
                        </ul>
                    </div>
                    <div class="content">
                    <p>Anexo: Arquivo Excel com detalhes dos aparelhos inativos</p>
                    </div>
                    <div class="footer">
                        Mensagem de envio automático, por favor não responda a mesma.
                    </div>
                    </body>
                    </html>
                    """

                    mensagem = MIMEMultipart()
                    mensagem['De'] = de
                    mensagem['Para'] = para
                    mensagem['Subject'] = assunto

                    # Anexar o arquivo Excel ao objeto mensagem existente
                    anexo_excel = MIMEBase('application', 'vnd.ms-excel')
                    with open(arquivo_relatorio, 'rb') as anexo:
                        anexo_excel.set_payload(anexo.read())
                    encoders.encode_base64(anexo_excel)
                    anexo_excel.add_header('Content-Disposition', f'attachment; filename="{arquivo_relatorio}"')
                    mensagem.attach(anexo_excel)

                    mensagem.attach(MIMEText(corpo_relatorio, 'html'))

                    texto_do_registro = mensagem.as_string()
                    server.sendmail(de, para, texto_do_registro)
                    print(f'Relatório enviado com sucesso para {para}')

                    # Registro de sucesso no log
                    registro_log = f'Relatório enviado com sucesso para {para}'
                    registrar_log(registro_log)

                    # Excluir o arquivo 'aparelhos_inativos.xlsx' após a conclusão do programa
                    if os.path.exists(arquivo_relatorio):
                        os.remove(arquivo_relatorio)
                        print("Arquivo 'aparelhos_inativos.xlsx' excluído com sucesso.")
                    else:
                        print("O arquivo 'aparelhos_inativos.xlsx' não existe.")

                except Exception as e:
                    print(f'Erro ao enviar o relatório por e-mail: {str(e)}')

                    # Registro de falha no log
                    registro_log = f'Erro ao enviar o relatório por e-mail: {str(e)}'
                    
                    registrar_log(registro_log)
                finally:
                    server.quit()

                registrar_log("Programa concluído.")
            self.enviado_com_sucesso.emit()

            # Excluir o arquivo 'dados_api.json' após a conclusão do programa
            if os.path.exists(arquivo_dados):
                os.remove(arquivo_dados)
                print("Arquivo 'dados_api.json' excluído com sucesso.")
            else:
                print("O arquivo 'dados_api.json' não existe.")
                
        except subprocess.CalledProcessError as e:
            # Se ocorrer um erro, emitir sinal de falha e passar a mensagem de erro
            self.falha_no_envio.emit("Erro durante o envio.")
        except Exception as ex:
            # Emitir sinal de falha e passar a mensagem de erro genérica
            self.falha_no_envio.emit(f"Erro durante o envio: {str(ex)}")
        finally:
            # Emitir sinal de finalização, independentemente do resultado
            self.finished.emit()

class Loading(QWidget):
    def __init__(self):
        super().__init__()

        icon_pixmap = QPixmap(os.path.join(os.path.dirname(__file__), "resources/titulo.png"))
        self.setWindowTitle("Loading")
        self.setGeometry(100, 100, 390, 340)
        self.setWindowIcon(QIcon(icon_pixmap))
        self.setWindowTitle("LOGMOBO")

        # Adicionando um QLabel para exibir o gif
        self.label = QLabel(self)
        self.label.setAlignment(Qt.AlignCenter)
        self.set_gif(os.path.join(os.path.dirname(__file__), "resources/loading.gif"))

        # Criação da instância de EnvioThread
        self.envio_thread = EnvioThread()

        # Conectar sinais após a criação da instância
        self.conectar_sinais()

        # Iniciar a thread para execução do envio.py
        self.iniciar_envio()

        # Layout vertical para organizar os elementos
        layout = QVBoxLayout()
        layout.addWidget(self.label)
        self.setLayout(layout)

    def set_gif(self, gif_path):
        # Configura o gif na QLabel
        movie = QMovie(gif_path)
        self.label.setMovie(movie)
        movie.start()
    
    def conectar_sinais(self):
        self.envio_thread.enviado_com_sucesso.connect(self.exibir_aviso_sucesso)
        self.envio_thread.falha_no_envio.connect(self.exibir_aviso_falha)
        self.envio_thread.finished.connect(self.finalizar_envio)

    def iniciar_envio(self):
        self.envio_thread.start()
    
    def centralizar_na_tela(self):
        screen_geometry = QDesktopWidget().screenGeometry()
        x = (screen_geometry.width() - self.width()) // 2
        y = (screen_geometry.height() - self.height()) // 2
        self.move(x, y)

    def exibir_aviso_sucesso(self):
        QMessageBox.information(self, "Sucesso", "Envio concluído com sucesso!")

    def exibir_aviso_falha(self, mensagem):
        QMessageBox.critical(self, "Falha no Envio", f"Erro ao enviar: {mensagem}")
        # Aguardar a interação do usuário antes de fechar a aplicação
        self.aguardar_interacao_usuario()

    def aguardar_interacao_usuario(self):
        loop = QEventLoop()
        QTimer.singleShot(4000, loop.quit)  # Aguardar 5 segundos
        loop.exec_()

        # Fechar a janela após o término do loop
        self.close()

    def finalizar_envio(self):
        # Finalizar a animação ou realizar qualquer ação necessária após o envio.py
        # Iniciar o loop de eventos para aguardar a interação do usuário
        self.aguardar_interacao_usuario()

if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Crie a instância da introdução e mostre-a
    intro_dialog = Introducao()
    intro_dialog.show()

    # Inicie o loop de eventos da aplicação
    app.exec_()

    # Após o encerramento da introdução, crie a instância da janela principal e mostre-a
    main_window = MainWindow()
    main_window.show()

    sys.exit(app.exec_())