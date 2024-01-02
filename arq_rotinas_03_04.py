import os
import logging
from datetime import datetime
import fitz  # PyMuPDF

def criar_log(nome, unidade, arquivo, data_hora):
    """
    Cria um registro no log.

    :param nome: Nome do paciente.
    :param unidade: Unidade de internação.
    :param arquivo: Nome do arquivo.
    :param data_hora: Data e hora do processo.
    """
    log_info = f"Paciente: {nome}, Unidade de Internação: {unidade}, Arquivo: {arquivo}, Data e Hora do Processo: {data_hora}"
    logging.info(log_info)

def obter_extensao(arquivo):
    """
    Obtém a extensão do arquivo em letras minúsculas.

    :param arquivo: Nome do arquivo.
    :return: Extensão do arquivo.
    """
    return os.path.splitext(arquivo)[1].lower()

def copiar_arquivo(arquivo_origem, destino):
    """
    Copia o arquivo de origem para o destino.

    :param arquivo_origem: Caminho do arquivo de origem.
    :param destino: Diretório de destino.
    """
    os.system(fr"copy {arquivo_origem} {destino}")

def renomear_arquivo(arquivo_origem, destino):
    """
    Renomeia o arquivo de origem para o destino.

    :param arquivo_origem: Caminho do arquivo de origem.
    :param destino: Novo nome ou caminho de destino.
    """
    os.rename(arquivo_origem, destino)

def compactar_pdf(arquivo_origem, destino):
    """
    Compacta o arquivo PDF de origem e salva no destino.

    :param arquivo_origem: Caminho do arquivo PDF de origem.
    :param destino: Caminho de destino para o PDF compactado.
    """
    with fitz.open(arquivo_origem) as pdf_document:
        pdf_document.save(destino, deflate=True)

def abrir_arquivo(arquivo):
    """
    Abre o arquivo especificado.

    :param arquivo: Caminho do arquivo a ser aberto.
    """
    os.startfile(arquivo)

def main():
    # Configurar o logger
    logging.basicConfig(filename=r'\\192.168.254.5\enfermagem\logs\log.txt', level=logging.INFO)

    # Obter informações do paciente
    nome_paciente = input("Digite o nome do paciente: ")
    unidade_internacao = input("Digite a unidade de internação: ")

    # Nome do arquivo
    nome_arquivo = f"{nome_paciente}_{unidade_internacao}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}"

    # Caminho do arquivo modelo
    caminho_modelo = fr"\\192.168.254.5\enfermagem"

    # Caminho para salvar os novos arquivos
    caminho_atual = fr"\\192.168.254.5\enfermagem\atuais"

    # Lista de arquivos para escolher
    arquivos = [
        "ANOTACAO_DO_TECNICO_MANUAL.docx",
        "EVOLUCAO_TEC_ENFERMAGEM_ENFERMARIA.doc",
        "HBV_SAE.docx",
        "EVOLUCAO_DE_ENFERMEIRO_UTI.pdf"
    ]

    # Menu de escolha
    while True:
        print("Escolha o arquivo a ser aberto:")
        for i, arquivo in enumerate(arquivos, 1):
            print(f"{i}. {arquivo}")
        print("0. Sair")

        opcao = input("Digite o número correspondente ao arquivo desejado: ")

        if opcao == "0":
            print("Saindo do programa...")
            break

        try:
            opcao = int(opcao)
            if opcao < 1 or opcao > len(arquivos):
                raise ValueError
        except ValueError:
            print("Opção inválida. Digite um número correspondente à opção desejada.")
            continue

        # Nome do arquivo escolhido
        arquivo_escolhido = arquivos[opcao - 1]

        # Caminho completo do arquivo escolhido
        caminho_arquivo_escolhido = fr"{caminho_modelo}\{arquivo_escolhido}"

        # Copiar o arquivo modelo para a pasta de arquivos atuais
        copiar_arquivo(caminho_arquivo_escolhido, fr"{caminho_atual}\{nome_arquivo}{obter_extensao(arquivo_escolhido)}")

        # Renomear o arquivo na pasta de arquivos atuais
        novo_caminho = fr"{caminho_atual}\{nome_paciente}_{unidade_internacao}_{arquivo_escolhido}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}{obter_extensao(arquivo_escolhido)}"

        # Compactar o PDF usando PyMuPDF se for um arquivo PDF
        if obter_extensao(arquivo_escolhido) == ".pdf":
            compactar_pdf(fr"{caminho_atual}\{nome_arquivo}{obter_extensao(arquivo_escolhido)}", novo_caminho)

            # Remover o arquivo original não compactado
            os.remove(fr"{caminho_atual}\{nome_arquivo}{obter_extensao(arquivo_escolhido)}")
        else:
            # Renomear diretamente para outros tipos de arquivo
            renomear_arquivo(fr"{caminho_atual}\{nome_arquivo}{obter_extensao(arquivo_escolhido)}", novo_caminho)

        
        # Abrir o arquivo
        abrir_arquivo(novo_caminho)

        # Criar log
        criar_log(nome_paciente, unidade_internacao, arquivo_escolhido, datetime.now().strftime('%Y-%m-%d %H:%M:%S'))

if __name__ == "__main__":
    main()
