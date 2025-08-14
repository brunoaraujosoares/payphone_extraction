#!/usr/bin/env python
# coding: utf-8

# # Converter PDFs para excel
# 
# O diretório **PDFs** contém uma lista com pastas com as siglas das UFs brasileiras. Dentro de cada subdiretório há um número desconhecido de arquivos em formato PDF.
# 
# O tamanho dos arquivos compactados é 56 GB
# 
# A sua missão, se você aceitar, é extrair os dados de 55 pontos de interesse dos PDFs e transofrmar em colunas de Excel. Para isso, será necessário 
# 
# - listar todos os pdfs dentro de cada diretório
#      - recuperar o nome do arquivo
#      - extrair o texto de cada campo de interesse usando expressões regulares
# - considerando a quantidade de arquivos, é melhor criar um arquivo CSV para cada diretório.

# # Importação de Bibliotecas

# In[1]:


import pandas as pd
from pathlib import Path
import fitz  # PyMuPDF
import re

from pdf2image import convert_from_path
import pytesseract
import cv2


# # Listar arquivos dos diretórios

# In[2]:


pasta_base = Path(r'C:\Users\bsoares\OneDrive - ANATEL\_atividades\2025\001-OI-analise-das-informacoes-de-continuidade\converter-pdfs\PDFs\')
arquivos_pdf = list( pasta_base.rglob("*.pdf") )


# ## Inicializando as Listas (colunas do dataframe)

# In[3]:


nome_arquivo = []
terminal_1_migrado = []
terminal_2_migrado = []
localidade = []
municipio = []
uf = []
latitude  = []
longitude  = []
status_terminal_pos_migracao  = []
visita_produtiva  = []
existe_cobertura_movel  = []
operadora  = []
tipo_instalacao = []
rede_linha_audio  = []
energia_utilizada = []
status_tup_antigo = []
deslocamento_fluvial = []
receptividade_responsavel = []
observacoes = []
data_implantacao = []
responsavel_implantacao = []

teste_lig_local_fixo_numero = []
teste_lig_local_fixo_resultado = []
teste_lig_a_cobrar_local_numero = []
teste_lig_a_cobrar_local_resultado = []
teste_lig_ddd_csp31_numero = []
teste_lig_ddd_csp31_resultado = []
teste_lig_ddd_csp14_numero = []
teste_lig_ddd_csp14_resultado = []
teste_lig_ddd_csp21_numero = []
teste_lig_ddd_csp21_resultado = []
teste_lig_celular_numero = []
teste_lig_celular_resultado = []
teste_lig_ddd_a_cobrar_csp31_numero = []
teste_lig_ddd_a_cobrar_csp31_resultado = []
teste_lig_ddd_a_cobrar_csp14_numero = []
teste_lig_ddd_a_cobrar_csp14_resultado = []
teste_lig_0800_numero = []
teste_lig_0800_resultado = []
teste_ldi_csp31_numero = []
teste_ldi_csp31_resultado = []
teste_ldi_csp14_numero = []
teste_ldi_csp14_resultado = []
teste_policia_numero = []
teste_policia_resultado = []
teste_bombeiro_numero = []
teste_bombeiro_resultado = []
teste_samu_numero = []
teste_samu_resultado = []
teste_chamada_recebida_local_fixo_numero = []
teste_chamada_recebida_local_fixo_resultado = []
teste_chamada_recebida_celular_numero = []
teste_chamada_recebida_celular_resultado = []
teste_chamada_recebida_a_cobrar_local_numero = []
teste_chamada_recebida_a_cobrar_local_resultado = []


# # funções de captura de texto

# In[4]:


def get_texto(texto: str, padrao: str) -> str | None:
    """
    Extrai do texto o conteúdo que combina com o padrão (RegEx).
    Retorna como string, ou None se o padrão não for encontrado.
    """
    resultado = re.search(padrao, texto)

    if resultado:
        return resultado.group(1)
    else:
        return None

def safe_get_texto(padrao):
    try:
        return '"' + get_texto(texto, padrao) + '"'
    except Exception:
        return None

def extract_text_from_image(image):
    text = pytesseract.image_to_string(image, lang= 'por')
    return text



# # Abrir o arquivo, extrair o texto e atribuir às variáveis 

# In[5]:


for pdf in arquivos_pdf:

    doc = fitz.open(pdf)
    texto = ""
    for pagina in doc:
        texto+= pagina.get_text()


    if not texto.strip():  # texto vazio ou só espaços
        # Tesseract
        pytesseract.pytesseract.tesseract_cmd = r"C:\Users\bsoares\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
        pages = convert_from_path(pdf)
        # Create a list to store extracted text from all pages
        texto = ""

        for page in pages:

            # Step 3: Extract text using OCR
            extract_text = extract_text_from_image(page)
            texto += extract_text


    caminho = Path(pdf)
    nome_arquivo.append( caminho.name )

    if caminho.name == 'COLR - 9636221050 - AP.pdf':
        print(texto)

    terminal_1_migrado.append( get_texto( texto, r"Terminal 1 Migrado:\s*([0-9]+)" ) )
    terminal_2_migrado.append( get_texto( texto, r"Terminal 2 Migrado:\s*([0-9]+)" ) )
    localidade.append( get_texto( texto,  r"Localidade:\s*(.*?)\n" ) )
    municipio.append( get_texto( texto , r"Município:\s*([\s\S]*?)\s*(?=\s*UF\s*:)") )
    uf.append( get_texto( texto , r"UF:[ \n](.*?)\n") )


    latitude.append(safe_get_texto(r"Latitude:[\n ](.*?)[ \n]")) 
    longitude.append(safe_get_texto(r"Longitude:\s(.*?)\w\n"))

    status_terminal_pos_migracao.append( get_texto( texto, r"Status do Terminal Após a Migração\s(.*?)\n" ) )
    visita_produtiva.append( get_texto( texto, r"Foi uma visita produtiva \(Com instalação dos equipamentos\)\?[ \n](.*?)\n" ) )
    existe_cobertura_movel.append( get_texto( texto, "(Sim|Não)[ \n]Operadora" ) )
    operadora.append( get_texto( texto, r"Operadora\?[ \n](.*?)\n" ) )   
    tipo_instalacao.append( get_texto( texto, r"Tipo de Instalação:[ \n]{1,2}(.*?)\n" ) )
    rede_linha_audio.append( get_texto( texto, r"Rede \/ Linha de Audio[ \n]{1,2}(.*?)\n" ) )
    energia_utilizada.append( get_texto( texto, r"Energia Utilizada:[ \n]{1,2}(.*?)\n" ) )    
    status_tup_antigo.append( get_texto( texto,  r"Status do TUP Antigo[ \n](.*?)\n" ) )
    deslocamento_fluvial.append( get_texto( texto, r"Deslocamento Fluvial Dedicado[ \n](.*?)\n" ) )
    observacoes.append( get_texto( texto, r"De[s]?crição Livre e Observações:[\n]{1,2}(.*?)\n" ) )
    data_implantacao.append( get_texto( texto,  r"Data da Implantação:[ \n]{1,2}(.*?)\n" ) )
    receptividade_responsavel.append( get_texto( texto , r"Como Foi a Receptividade do Responsável Pelo\nNovo Ponto de Instalação Para a Nova Solução Oi[ \n]{1,2}(.*?)\n") )
    responsavel_implantacao.append( get_texto ( texto, r"Responsável[\n]Implantação:[ \n]{1,2}(.*?)\n" ) ) 
    teste_lig_local_fixo_numero.append( get_texto( texto, r"Ligação Local Fixo[ \n]([0-9 \-]*)[ \n][A-Za-z]" ) )
    teste_lig_local_fixo_resultado.append( get_texto( texto, r"Ligação Local Fixo[ \n][0-9 \- \n]*(.*?)\n" ) )
    teste_lig_a_cobrar_local_numero.append( get_texto( texto, r"Ligação a Cobrar Local[ \n]([0-9 \-]*)\W" ) )
    teste_lig_a_cobrar_local_resultado.append( get_texto( texto, r"Ligação a Cobrar Local[ \n][0-9 \-]*[ \n](.*?)\n" ) )
    teste_lig_ddd_csp31_numero.append( get_texto( texto, r"Ligação DDD com CSP 31[ \n]([0-9 \-]*)" ).strip() )
    teste_lig_ddd_csp31_resultado.append( get_texto( texto, r"Ligação DDD com CSP 31[ \n][0-9 \- \n]*(.*?)\n" ) )
    teste_lig_ddd_csp14_numero.append( get_texto( texto, r"Ligação DDD com CSP 14[ \n]*([N\/A0-9 \-]*)" ).strip() ) 
    # teste_lig_ddd_csp14_resultado.append( get_texto( texto, r"Ligação DDD com CSP 14[ \n]*[0-9 \- \n]*(.*?)\n" ) )
    teste_lig_ddd_csp14_resultado.append( get_texto( texto, r"Ligação DDD com CSP 14[ \n]*(?:N\/A|[\d\s\-\n]*)(.*?)\n" ) )
    teste_lig_ddd_csp21_numero.append( get_texto( texto, r"Ligação DDD CSP 21[ \n]([0-9 \-]*)" ).strip() )
    teste_lig_ddd_csp21_resultado.append( get_texto( texto, r"Ligação DDD CSP 21[ \n][0-9 \- \n]*(.*?)\n" ) )
    teste_lig_celular_numero.append( get_texto( texto, r"Ligação Celular[ \n]([N\/A0-9 \-]*)" ).strip() ) 
    #teste_lig_celular_resultado.append( get_texto ( texto,  r"Ligação Celular[ \n][0-9 \- \n]*(.*?)\n" ) )    
    teste_lig_celular_resultado.append( get_texto ( texto,  r"Ligação Celular[ \n](?:N\/A|[\d\s\-\n]*)(.*?)\n" ) )    

    teste_lig_ddd_a_cobrar_csp31_numero.append( get_texto( texto, r"Ligação DDD à cobrar CSP 31[ \n]([0-9 \-]*)" ).strip() )

    teste_lig_ddd_a_cobrar_csp31_resultado.append( get_texto( texto , r"Ligação DDD à cobrar CSP 31[ \n0-9 \-]*[ \n]([A-Za-z ].*)") )
    teste_lig_ddd_a_cobrar_csp14_numero.append( get_texto( texto, r"Ligação DDD à cobrar CSP 14[ \n]([0-9 \- N\/A]*)[ \n]" ) )
    teste_lig_ddd_a_cobrar_csp14_resultado.append( get_texto( texto , r"Ligação DDD à cobrar CSP 14[ \n0-9 \-]*([A-Za-z ].*)") )
    teste_lig_0800_numero.append( get_texto( texto, r"Ligação 0800[ \n]([0-9 ]*)" ).strip() )
    teste_lig_0800_resultado.append( get_texto( texto, r"Ligação 0800[ \n0-9 ]*([A-Za-z ].*)" ) )    
    teste_ldi_csp31_numero.append( get_texto( texto, r"LDI CSP 31 - Internacional[ \n]([0-9 \-]*)[ \nA-Za-z]" ) )
    teste_ldi_csp31_resultado.append( get_texto( texto, r"LDI CSP 31 - Internacional[ \n 0-9 \-]*([A-Za-z].*)" ) )
    teste_ldi_csp14_numero.append( get_texto( texto,  r"LDI CSP 14 - Internacional[ \n]([0-9 \-]*)[A-Za-z ]*" ) )
    teste_ldi_csp14_resultado.append( get_texto( texto, r"LDI CSP 14 - Internacional[\n[0-9 \-]*([A-Za-z ].*)" ).strip() )
    teste_policia_numero.append( get_texto( texto, r"Polícia[ \n]([0-9 \-]*)[ \nA-Za-z]*").strip() )
    teste_policia_resultado.append( get_texto( texto, r"Polícia[ \n0-9 \-]*([ \nA-Za-z].*)").strip() )
    teste_bombeiro_numero.append( get_texto( texto, r"Bombeiro[ \n]([0-9 \-]*)[ \nA-Za-z]*").strip() )
    teste_bombeiro_resultado.append( get_texto( texto, r"Bombeiro[ \n0-9 \-]*([ \nA-Za-z].*)").strip() )
    teste_samu_numero.append( get_texto( texto , r"SAMU[ \n]([0-9 \-]*)[ \n]" ) )          
    teste_samu_resultado.append( get_texto( texto,  r"SAMU[ \n0-9 \-]*([ A-Za-z].*)\n" ) )

    teste_chamada_recebida_local_fixo_numero.append( get_texto( texto, r"Chamada Recebida Local Fixo[ \n]([0-9 \-]*)[ \n]" ) )
    teste_chamada_recebida_local_fixo_resultado.append( get_texto( texto , r"Chamada Recebida Local Fixo[ \n 0-9\-]*([A-Za-z ].*)") )
    teste_chamada_recebida_celular_numero.append( get_texto( texto, r"Chamada Recebida Celular[ \n]([0-9 \-]*)[ \n]"  ) )
    teste_chamada_recebida_celular_resultado.append( get_texto( texto, r"Chamada Recebida Celular[ \n 0-9\-]*([A-Za-z ].*)" ) )
    teste_chamada_recebida_a_cobrar_local_numero.append( get_texto( texto, r"Chamada Recebida a Cobrar Local[ \n]([0-9 \-]*)"  ).strip() )    
    teste_chamada_recebida_a_cobrar_local_resultado.append( get_texto( texto, r"Chamada Recebida a Cobrar Local[ \n0-9 \-]*([ A-Za-z].*)" ) )


    doc.close()


# # criando o dataframe

# ## variaveis

# In[6]:


# colunas do dataframe
colunas = [
    "nome_arquivo",
    "terminal_1_migrado",
    "terminal_2_migrado",
    "localidade",
    "municipio",
    "uf",
    "latitude",
    "longitude",
    "status_terminal_pos_migracao",
    "visita_produtiva",
    "existe_cobertura_movel",
    "operadora",
    "tipo_instalacao",
    "rede_linha_audio",
    "energia_utilizada",
    "status_tup_antigo",
    "deslocamento_fluvial",
    "receptividade_responsavel",
    "observacoes",
    "data_implantacao",
    "responsavel_implantacao",

    # Testes de ligações (número + resultado)
    "teste_lig_local_fixo_numero",
    "teste_lig_local_fixo_resultado",
    "teste_lig_a_cobrar_local_numero",
    "teste_lig_a_cobrar_local_resultado",
    "teste_lig_ddd_csp31_numero",
    "teste_lig_ddd_csp31_resultado",
    "teste_lig_ddd_csp14_numero",
    "teste_lig_ddd_csp14_resultado",
    "teste_lig_ddd_csp21_numero",
    "teste_lig_ddd_csp21_resultado",
    "teste_lig_celular_numero",
    "teste_lig_celular_resultado",
    "teste_lig_ddd_a_cobrar_csp31_numero",
    "teste_lig_ddd_a_cobrar_csp31_resultado",
    "teste_lig_ddd_a_cobrar_csp14_numero",
    "teste_lig_ddd_a_cobrar_csp14_resultado",
    "teste_lig_0800_numero",
    "teste_lig_0800_resultado",
    "teste_ldi_csp31_numero",
    "teste_ldi_csp31_resultado",
    "teste_ldi_csp14_numero",
    "teste_ldi_csp14_resultado",
    "teste_policia_numero",
    "teste_policia_resultado",
    "teste_bombeiro_numero",
    "teste_bombeiro_resultado",
    "teste_samu_numero",
    "teste_samu_resultado",
    "teste_chamada_recebida_local_fixo_numero",
    "teste_chamada_recebida_local_fixo_resultado",
    "teste_chamada_recebida_celular_numero",
    "teste_chamada_recebida_celular_resultado",
    "teste_chamada_recebida_a_cobrar_local_numero",
    "teste_chamada_recebida_a_cobrar_local_resultado"
]


# In[7]:


dados = {coluna: globals()[coluna] for coluna in colunas}


# In[8]:


## criar o dataframe com as colulas de interesse
df_tups = pd.DataFrame(dados, dtype=str)


# In[9]:


df_tups.to_csv("tups.csv", index=False, sep=";", encoding="ISO-8859-1")

