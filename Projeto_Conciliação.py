"""
Projeto para ajudar o financeiro copiar os dados da planilha de extrato e passar para a planilha de transferencia para
poder importar o arquivo csv no LBC.
Esse projeto está tratando a parte de cartões.
"""

import concurrent
import os
import csv
import sys
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta
import xlrd
from openpyxl import load_workbook, Workbook
import pandas as pd
from bs4 import BeautifulSoup
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QPushButton, QLabel, QLineEdit, QWidget, QProgressBar, QMessageBox, QDialog
)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QIcon

# Função para criar a pasta Conciliação no Disco C
def criar_pasta_conciliacao():
    pasta_conciliacao = r"C:\Conciliação"
    if not os.path.exists(pasta_conciliacao):
        os.makedirs(pasta_conciliacao)
        print(f"Pasta Conciliação criada em {pasta_conciliacao}")
    return pasta_conciliacao

# Função para converter arquivos .xls para .xlsx
def convert_single_xls(xls_path):
    xlsx_path = xls_path.replace(".xls", ".xlsx")
    try:
        try:
            df = pd.read_excel(xls_path, engine="xlrd")
            df.to_excel(xlsx_path, index=False, engine="openpyxl")
        except Exception:
            tables = pd.read_html(xls_path)
            with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
                for i, table in enumerate(tables):
                    table.to_excel(writer, sheet_name=f"Sheet{i + 1}", index=False)
        print(f'Convertido: {os.path.basename(xls_path)} -> {os.path.basename(xlsx_path)}')
    except Exception as e:
        print(f'Erro ao converter {os.path.basename(xls_path)}: {e}')

def convert_xls_to_xlsx(folder_path):
    xls_files = [os.path.join(folder_path, file) for file in os.listdir(folder_path) if file.endswith(".xls")]

    with concurrent.futures.ThreadPoolExecutor(max_workers=os.cpu_count()) as executor:
        executor.map(convert_single_xls, xls_files)

# Função para formatar a data com ano específico
def formatar_data(data_str, ano="2025"):
    try:
        data = datetime.strptime(f"{data_str}/{ano}", "%d/%m/%Y")
        return data.strftime("%d/%m/%Y")  # Ou "%d/%m/%Y" para ano com 4 dígitos
    except ValueError:
        return ""

def processar_sispag_fornecedores(sheet_destino, proxima_linha, coluna_b, coluna_e, coluna_f, mapeamento_base_itau, chave_busca):
    # Coluna A: Adiciona os dados da conta que foi encontrada na coluna D da sheet base itau.
    if chave_busca in mapeamento_base_itau:
        coluna_a = mapeamento_base_itau[chave_busca][1]  # Coluna D é o segundo elemento da tupla
    else:
        coluna_a = ""  # Se não houver dados, deixe vazio

    # Coluna F: Será a seguinte conta fixa: 02 - ITAU 35422-2 CC
    coluna_f_fixa = "02 - ITAU 35422-2 CC"

    # Coluna G será fixa com o ID 201
    coluna_g_fixa = 201

    # Verifica se o valor é negativo e converte para positivo
    try:
        coluna_f = float(coluna_f)  # Certifica que o valor é numerico
        if "SISPAG FORNECEDORES" in str(coluna_e) and coluna_f < 0:
            coluna_f = abs(coluna_f)  # Remove o sinal negativo
    except ValueError:
        pass  # Caso o valor não seja numerico, manter como está

    # Adiciona a nova linha na planilha de destino
    sheet_destino.append([
        coluna_a,  # Coluna A
        mapeamento_base_itau.get(chave_busca, ('', ''))[0],  # Coluna B
        coluna_b,  # Coluna C (data formatada)
        coluna_e,  # Coluna D
        coluna_f,  # Coluna E
        coluna_f_fixa,  # Coluna F
        coluna_g_fixa  # Coluna G
    ])

    return proxima_linha + 1

def processar_dev_pix(sheet_destino, proxima_linha, coluna_b, coluna_e, coluna_f, mapeamento_base_itau, chave_busca):
    # Coluna A: Recebe o valor que estava na coluna F (variável da base)
    if chave_busca in mapeamento_base_itau:
        coluna_a = mapeamento_base_itau[chave_busca][1]  # Coluna D é o segundo elemento da tupla
    else:
        coluna_a = ""  # Se não houver dados, deixe vazio

    # Coluna F: Recebe o valor que estava na coluna A ("04 - PIX")
    coluna_f_fixa = "04 - PIX"

    # Coluna G: recebe o ID da base
    coluna_g = mapeamento_base_itau.get(chave_busca, ('', ''))[0]

    # Verifica se o valor é negativo e converte para positivo
    try:
        coluna_f = float(coluna_f)  # Certifica que o valor é numérico
        if "DEV PIX" in str(coluna_e) and coluna_f < 0:
            coluna_f = abs(coluna_f)  # Remove o sinal negativo
    except ValueError:
        pass  # Caso o valor não seja numérico, manter como está

    # Adiciona a nova linha na planilha de destino
    sheet_destino.append([
        coluna_a,  # Coluna A (valor da base)
        mapeamento_base_itau.get(chave_busca, ('', ''))[0],  # Coluna B
        coluna_b,  # Coluna C (data formatada)
        coluna_e,  # Coluna D
        coluna_f,  # Coluna E (valor positivo)
        coluna_f_fixa,  # Coluna F ("04 - PIX")
        coluna_g  # Coluna G
    ])

    return proxima_linha + 1

def processar_todas_informacoes(caminho_extrato, book_destino, caminho_base, datas_validas):
    try:
        # Nome das abas
        sheet_extrato = "Sheet1"
        sheet_destino = "TRANSFERENCIAS"
        sheet_base_cartoes = "BASE CARTOES"
        sheet_base_itau = "BASE ITAU"

        print(f"Abrindo arquivo de extrato: {caminho_extrato}")
        book_extrato = load_workbook(caminho_extrato)
        sheet_extrato = book_extrato[sheet_extrato]

        print(f"Selecionando planilha de destino: {sheet_destino}")
        sheet_destino = book_destino[sheet_destino]

        print(f"Abrindo arquivo de base: {caminho_base}")
        book_base = load_workbook(caminho_base)
        sheet_base_cartoes = book_base[sheet_base_cartoes]
        sheet_base_itau = book_base[sheet_base_itau]

        # Encontrar a próxima linha vazia na planilha de destino
        proxima_linha = sheet_destino.max_row + 1
        print(f"Próxima linha vazia na planilha de destino: {proxima_linha}")

        # Ler todos os tipos de cartão, PIX e Shell Box da base de cartões
        tipos_cartoes = set()
        tipos_pix = set()
        tipos_shellbox = set()

        for row in sheet_base_cartoes.iter_rows(min_row=2, values_only=True):
            tipo_cartao = row[0]  # Coluna A: Cartões
            tipo_pix = row[1]  # Coluna B: PIX
            tipo_shellbox = row[2]  # Coluna C: Shell Box

            if tipo_cartao:
                tipos_cartoes.add(tipo_cartao.strip())
            if tipo_pix:
                tipos_pix.add(tipo_pix.strip())
            if tipo_shellbox:
                tipos_shellbox.add(tipo_shellbox.strip())

        print(f"Tipos de cartões encontrados: {tipos_cartoes}")
        print(f"Tipos de PIX encontrados: {tipos_pix}")
        print(f"Tipos de Shell Box encontrados: {tipos_shellbox}")

        # Ler a chave de busca da planilha de extrato
        chave_busca = sheet_extrato['E6'].value
        print(f"Chave de busca encontrada: {chave_busca}")

        # Criar um dicionário para mapear as informações da planilha Base Itau
        mapeamento_base_itau = {}
        for row in sheet_base_itau.iter_rows(min_row=2, values_only=True):
            coluna_b = row[1]  # Coluna B
            coluna_c = row[2]  # Coluna C (chave de busca)
            coluna_d = row[3]  # Coluna D (dados para a coluna A)
            if coluna_c:
                mapeamento_base_itau[coluna_c] = (coluna_b, coluna_d)
        print(f"Mapeamento Base Itau criado com {len(mapeamento_base_itau)} entradas")

        # Percorrer todas as linhas da planilha de extrato
        for row in sheet_extrato.iter_rows(min_row=12, values_only=True):
            coluna_b = row[1]  # Data
            coluna_e = row[4]  # Descrição
            coluna_f = row[5]  # Valor

            if coluna_b:
                coluna_b = coluna_b.strip()
                coluna_b = coluna_b.replace('\xa0', '')

            if coluna_b in datas_validas and coluna_e:
                # Verificar se é "SISPAG FORNECEDORES"
                if "SISPAG FORNECEDORES" in str(coluna_e):
                    print(f"Processando SISPAG FORNECEDORES na linha: {row}")
                    data_formatada = formatar_data(coluna_b, ano="2025")  # Formatar a data com ano 2025
                    proxima_linha = processar_sispag_fornecedores(sheet_destino, proxima_linha, data_formatada, coluna_e, coluna_f, mapeamento_base_itau, chave_busca)
                # Verificar se é "DEV PIX"
                elif "DEV PIX" in str(coluna_e):
                    print(f"Processando DEV PIX na linha: {row}")
                    data_formatada = formatar_data(coluna_b, ano="2025")  # Formatar a data com ano 2025
                    proxima_linha = processar_dev_pix(sheet_destino, proxima_linha, data_formatada, coluna_e, coluna_f, mapeamento_base_itau, chave_busca)
                else:
                    # Processar outras informações (cartões, PIX, Shell Box)
                    processado = False

                    # Verificar se é um cartão
                    for tipo_cartao in tipos_cartoes:
                        if str(coluna_e).startswith(tipo_cartao):
                            data_formatada = formatar_data(coluna_b, ano="2025")  # Formatar a data com ano 2025
                            sheet_destino.append([
                                "04 - BAIXA CARTOES A RECEBER",  # Informação fixa para cartões
                                mapeamento_base_itau.get(chave_busca, ('', ''))[0],
                                data_formatada,
                                coluna_e,
                                coluna_f,
                                mapeamento_base_itau.get(chave_busca, ('', ''))[1],
                                mapeamento_base_itau.get(chave_busca, ('', ''))[0]
                            ])
                            proxima_linha += 1
                            processado = True
                            break  # Sair do loop após processar o tipo de cartão

                    # Verificar se é PIX
                    if not processado:
                        for tipo_pix in tipos_pix:
                            if str(coluna_e).startswith(tipo_pix):
                                data_formatada = formatar_data(coluna_b, ano="2025")  # Formatar a data com ano 2025
                                sheet_destino.append([
                                    "04 - PIX",  # Informação fixa para PIX
                                    mapeamento_base_itau.get(chave_busca, ('', ''))[0],
                                    data_formatada,
                                    coluna_e,
                                    coluna_f,
                                    mapeamento_base_itau.get(chave_busca, ('', ''))[1],
                                    mapeamento_base_itau.get(chave_busca, ('', ''))[0]
                                ])
                                proxima_linha += 1
                                processado = True
                                break  # Sair do loop após processar o tipo de PIX

                    # Verificar se é Shell Box
                    if not processado:
                        for tipo_shellbox in tipos_shellbox:
                            if str(coluna_e).startswith(tipo_shellbox):
                                data_formatada = formatar_data(coluna_b, ano="2025")  # Formatar a data com ano 2025
                                sheet_destino.append([
                                    "04 - BAIXA CARTOES SHELL BOX",  # Informação fixa para Shell Box
                                    mapeamento_base_itau.get(chave_busca, ('', ''))[0],
                                    data_formatada,
                                    coluna_e,
                                    coluna_f,
                                    mapeamento_base_itau.get(chave_busca, ('', ''))[1],
                                    mapeamento_base_itau.get(chave_busca, ('', ''))[0]
                                ])
                                proxima_linha += 1
                                processado = True
                                break  # Sair do loop após processar o tipo de Shell Box
        print("Processamento concluído com sucesso.")
    except Exception as e:
        print(f"Erro durante o processamento: {e}")


# Classe para executar o processamento em uma thread separada
class Worker(QThread):
    finished = pyqtSignal()

    def __init__(self, pasta_extratos, pasta_csv, caminho_destino, caminho_base, datas_validas):
        super().__init__()
        self.pasta_extratos = pasta_extratos
        self.pasta_csv = pasta_csv
        self.caminho_destino = caminho_destino
        self.caminho_base = caminho_base
        self.datas_validas = datas_validas

    def run(self):
        # Processar todos os arquivos de extrato na pasta Extratos
        for arquivo in os.listdir(self.pasta_extratos):
            if arquivo.endswith(".xlsx"):  # Processar apenas arquivos Excel
                caminho_extrato = os.path.join(self.pasta_extratos, arquivo)
                nome_arquivo_csv = arquivo.replace(".xlsx", ".csv")
                caminho_csv = os.path.join(self.pasta_csv, nome_arquivo_csv)

                print(f"\nProcessando arquivo: {arquivo}")

                # Abrir a planilha de destino apenas uma vez
                book_destino = load_workbook(self.caminho_destino)

                # Executar o processamento para todas as informações
                processar_todas_informacoes(caminho_extrato, book_destino, self.caminho_base, self.datas_validas)

                # Salvar a planilha de destino uma única vez
                book_destino.save(self.caminho_destino)

                # Salvar a planilha de destino como CSV com separador de vírgula
                with open(caminho_csv, mode="w", newline="", encoding="latin1") as file:
                    writer = csv.writer(file, delimiter=";")  # Usar ';' como separador
                    # Escrever os dados da planilha de destino no arquivo CSV
                    for row in book_destino["TRANSFERENCIAS"].iter_rows(values_only=True):
                        writer.writerow(row)

                print(f"Dados salvos em {caminho_csv}")

                # Limpar a planilha de destino mantendo apenas o cabeçalho
                sheet_destino = book_destino["TRANSFERENCIAS"]
                for row in sheet_destino.iter_rows(min_row=2):
                    for cell in row:
                        cell.value = None

                book_destino.save(self.caminho_destino)

        self.finished.emit()


# Interface gráfica
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Processamento de Extratos")
        self.setGeometry(200, 200, 550, 250)

        # Layout principal
        layout = QVBoxLayout()

        # Define o ícone da janela
        self.setWindowIcon(QIcon('contrato.ico'))

        # Botão para carregar arquivos
        self.btn_carregar = QPushButton("Carregar Arquivos (.xls para .xlsx)")
        self.btn_carregar.clicked.connect(self.carregar_arquivos)
        self.btn_carregar.setStyleSheet(
            "QPushButton {"
            "   border-radius: 15px;"
            "   font-size: 16px;"
            "   padding: 10px;"
            "   background-color: #329932;"
            "   color: white;"
            "}"
            "QPushButton:hover {"
            "   background-color: #006600;"
            "}"
        )
        layout.addWidget(self.btn_carregar)

        # Campo para inserir as datas
        self.label_datas = QLabel("Digite as datas desejadas (formato dd/mm, separadas por vírgula):")
        self.label_datas.setStyleSheet("font-size: 16px;")
        layout.addWidget(self.label_datas)

        self.input_datas = QLineEdit()
        self.input_datas.setStyleSheet(
            "QLineEdit {"
            "   font-size: 16px;"
            "   padding: 10px;"
            "   border: 2px solid #ccc;"
            "   border-radius: 10px;"
            "}"
        )
        layout.addWidget(self.input_datas)

        # Botão para iniciar o processamento
        self.btn_processar = QPushButton("Processar")
        self.btn_processar.clicked.connect(self.iniciar_processamento)
        self.btn_processar.setStyleSheet(
            "QPushButton {"
            "   border-radius: 15px;"
            "   font-size: 16px;"
            "   padding: 10px;"
            "   background-color: #4c4cff;"
            "   color: white;"
            "}"
            "QPushButton:hover {"
            "   background-color: #0000cc;"
            "}"
        )
        layout.addWidget(self.btn_processar)

        # Barra de progresso
        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)

        # Widget central
        widget = QWidget()
        widget.setLayout(layout)
        self.setCentralWidget(widget)

        # Criar a pasta Conciliação no Disco C
        self.pasta_conciliacao = criar_pasta_conciliacao()

        # Caminhos dos arquivos
        self.pasta_extratos = r"C:\Conciliação\Extratos"
        self.pasta_csv = r"C:\Conciliação\Arquivos_CSV"
        self.caminho_destino = r"C:\Conciliação\TRANSFERENCIAS.xlsx"
        self.caminho_base = r"C:\Conciliação\Base.xlsx"

        # Criar a pasta Arquivos_CSV se não existir
        if not os.path.exists(self.pasta_csv and self.pasta_extratos):
            os.makedirs(self.pasta_csv)
            os.makedirs(self.pasta_extratos)

    def carregar_arquivos(self):
        # Criar a janela de carregamento
        self.carregando_dialog = QDialog(self)
        self.carregando_dialog.setWindowTitle("Carregando Arquivos...")
        self.carregando_dialog.setFixedSize(300, 100)
        self.carregando_label = QLabel("Carregando e convertendo arquivos...", self.carregando_dialog)
        self.carregando_label.setAlignment(Qt.AlignCenter)
        self.progress_carregamento = QProgressBar(self.carregando_dialog)
        self.progress_carregamento.setRange(0, 1)  # Barra de progresso indeterminada
        layout_dialog = QVBoxLayout()
        layout_dialog.addWidget(self.carregando_label)
        layout_dialog.addWidget(self.progress_carregamento)
        self.carregando_dialog.setLayout(layout_dialog)

        # Exibir a janela de carregamento
        self.carregando_dialog.show()

        # Converter arquivos .xls para .xlsx
        convert_xls_to_xlsx(self.pasta_extratos)

        # Fechar a janela de carregamento e mostrar mensagem
        self.carregando_dialog.accept()
        QMessageBox.information(self, "Sucesso", "Arquivos convertidos com sucesso!")


    def iniciar_processamento(self):
        # Obter as datas inseridas pelo usuário
        datas_input = self.input_datas.text()
        if not datas_input:
            QMessageBox.warning(self, "Erro", "Por favor, insira as datas desejadas.")
            return

        # Converter a entrada do usuário em um conjunto de datas válidas
        datas_validas = set(data.strip() for data in datas_input.split(",") if data.strip())

        # Iniciar o processamento em uma thread separada
        self.worker = Worker(self.pasta_extratos, self.pasta_csv, self.caminho_destino, self.caminho_base, datas_validas)
        self.worker.finished.connect(self.processamento_concluido)
        self.worker.start()

        # Atualizar a interface
        self.progress_bar.setRange(0, 0)  # Barra de progresso indeterminada
        self.btn_processar.setEnabled(False)

    def processamento_concluido(self):
        # Finalizar a barra de progresso
        self.progress_bar.setRange(0, 1)
        self.progress_bar.setValue(1)

        # Habilitar o botão de processamento
        self.btn_processar.setEnabled(True)

        # Exibir mensagem de conclusão
        QMessageBox.information(self, "Concluído", "Processamento finalizado com sucesso!")

# Executar a aplicação
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


















