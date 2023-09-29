# PROJETO DE AUTOMAÇÃO DE TAREFAS #

# Importação das Bibliotecas usadas #
# Para usar cada uma é preciso baixar se for a primeira vez: 
    # pip3 install pyautoguiCAHA000273 openpyxl
import pyautogui
import time
import pandas as pd

# Principais comandos do PYAUTOGUI #
    # pyautogui.write -> escrever um texto
    # pyautogui.press -> apertar 1 tecla
    # pyautogui.click -> clicar em algum lugar da tela
    # pyautogui.hotkey -> combinação de teclas

                            # Passo 1: Entrar no sistema da empresa
                                # https://dlp.hashtagtreinamentos.com/python/intensivao/login

# Pausa padrão para cada etapa do código
pyautogui.PAUSE = 0.6

# Abrir o chrome
pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("enter")

# Entrar no link
link = "https://dlp.hashtagtreinamentos.com/python/intensivao/login"
pyautogui.write(link)
pyautogui.press("enter")

# Tempo de espera para o site carregar
time.sleep(5)

                                # Passo 2: Fazer login

# Clique no campo de login da página
# Dois parâmetros opcionais: 
    # 1ª clicks=2               quantidade de cliques
    # 2ª button="right"         botão direito ou esquerdo do mouse
pyautogui.click(x=2784, y=404)
pyautogui.write("teste@email.com")
pyautogui.press("tab")
pyautogui.write("senha123")
pyautogui.press("tab")
pyautogui.press("enter")

                                # Passo 3: Importar a base de dados de produtos

tabela = pd.read_csv("produtos.csv")

                                # Passo 4: Cadastrar produtos

# Como esses códigos precisam ser repetidos, inclui eles em um laço for
    # tabela.index significa que irá pegar as linhas | tabela.columns pega as colunas
for linha in tabela.index:

    # clicar no campo de código
    pyautogui.click(x=2755, y=284)

    # pegar da tabela o valor do campo que a gente quer preencher
    codigo = tabela.loc[linha, "codigo"]

    # preencher o campo
    pyautogui.write(str(codigo))
    # passar para o proximo campo
    pyautogui.press("tab")
    # preencher o campo
    pyautogui.write(str(tabela.loc[linha, "marca"]))
    pyautogui.press("tab")
    pyautogui.write(str(tabela.loc[linha, "tipo"]))
    pyautogui.press("tab")
    pyautogui.write(str(tabela.loc[linha, "categoria"]))
    pyautogui.press("tab")
    pyautogui.write(str(tabela.loc[linha, "preco_unitario"]))
    pyautogui.press("tab")
    pyautogui.write(str(tabela.loc[linha, "custo"]))
    pyautogui.press("tab")
    obs = tabela.loc[linha, "obs"]

    if not pd.isna(obs):
        pyautogui.write(str(tabela.loc[linha, "obs"]))

    pyautogui.press("tab")

    # Cadastra o produto (clicando no botão enviar)
    pyautogui.press("enter")

    # Usar o scroll de tudo para cima
    pyautogui.scroll(5000)

                                # Passo 5: Repetir o cadastro de todos os produtos


                                # Para interromper o programa clique no terminar e dê um ctrol + c 