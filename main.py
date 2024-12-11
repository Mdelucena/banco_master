from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import datetime, timedelta
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
from selenium.webdriver.common.action_chains import ActionChains
from dotenv import load_dotenv
import os

load_dotenv()
USUARIO = os.getenv("USUARIO_MASTER")
SENHA = os.getenv("SENHA_MASTER")
print(f"USUARIO: {USUARIO}")
print(f"SENHA: {SENHA}")
navegador = webdriver.Chrome()
navegador.get("https://autenticacao.bancomaster.com.br/login")
navegador.maximize_window()

# Preencher login
print("Preenchendo usuario...")
navegador.find_element(
    By.XPATH, '//*[@id="mat-input-0"]').send_keys(USUARIO)
print("Preenchendo senha...")
navegador.find_element(By.XPATH, '//*[@id="mat-input-1"]').send_keys(SENHA)

# Clicar no botão "Entrar"
print("Entrando...")
botao_entrar = navegador.find_element(
    By.XPATH, '/html/body/app-root/app-login/div/div[2]/mat-card/mat-card-content/form/div[3]/button[2]')
botao_entrar.click()

# Tratamento para a tela de "já conectado"
try:
    botao_sim = WebDriverWait(navegador, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="mat-dialog-0"]/app-confirmacao-dialog/div/div[3]/div/app-botao-icon-v2[2]/button'))
    )
    botao_sim.click()
    print("Conexão anterior confirmada com 'Sim'.")
except:
    print("Tela de conexão anterior não apareceu.")


# Clicar no item menu_tetris
menu_tetris = WebDriverWait(navegador, 10).until(
    EC.element_to_be_clickable(
        (By.XPATH, "//i[contains(@class, 'icon-element-3')]"))
)
print("clicando no tetris...")
menu_tetris.click()

# Esperar que o modal ou o conteúdo relacionado tenha sido carregado completamente
# Isso pode incluir esperar até que o modal ou um elemento específico dentro dele esteja visível
time.sleep(2)
modal_esperado = WebDriverWait(navegador, 10).until(
    # Verifica se o elemento do modal está visível
    EC.visibility_of_element_located(
        (By.XPATH, '//*[@id="dorp_list_icons_0"]'))
)
print("Esperando modal...")
# Agora, clicar no elemento dentro do modal
vendas_consig = WebDriverWait(navegador, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="dorp_list_icons_0"]'))
)
print("clicando em vendas consignadas...")
vendas_consig.click()

time.sleep(5)
menu_hamburguer = WebDriverWait(navegador, 10).until(
    EC.element_to_be_clickable(
        (By.XPATH, '/html/body/app-root/app-home-v2/app-master-header-layout/mat-toolbar/button/span[1]/mat-icon'))
)
print("Apertando menu hamburguer...")
menu_hamburguer.click()

time.sleep(2)

cadastro_saque = WebDriverWait(navegador, 10).until(
    EC.element_to_be_clickable(
        (By.XPATH, '//*[@id="mat-expansion-panel-header-0"]/span/mat-panel-title'))
)
print("clicando cadastro saque...")
cadastro_saque.click()

botao_consul_esteira = WebDriverWait(navegador, 10).until(
    EC.element_to_be_clickable(
        (By.XPATH, '//*[@id="cdk-accordion-child-0"]/div/mat-list-item[2]/span/a/span[1]'))
)
print("clicando botao consultar esteira...")
botao_consul_esteira.click()


# Obtém a data de hoje (data final)
data_final = datetime.now()
data_inicial = data_final - timedelta(days=30)
data_inicial_formatada = data_inicial.strftime('%d/%m/%Y')
data_final_formatada = data_final.strftime('%d/%m/%Y')
campo_input_esquerda = WebDriverWait(navegador, 20).until(
    EC.visibility_of_element_located(
        (By.XPATH, "//input[contains(@class, 'mat-start-date')]"))
)

campo_input_esquerda.clear()
campo_input_esquerda.send_keys(data_inicial_formatada)


campo_input_direita = WebDriverWait(navegador, 20).until(
    EC.element_to_be_clickable(
        (By.XPATH, "//input[contains(@class, 'mat-end-date')]"))
)

campo_input_direita.clear()
campo_input_direita.send_keys(data_final_formatada)
print("Data selecionada...")


botao_status_forma = WebDriverWait(navegador, 5).until(
    EC.element_to_be_clickable(
        (By.XPATH, '//*[@id="mat-select-value-7"]/span'))
)
print("selecionando status forma...")
botao_status_forma.click()

status_forma = WebDriverWait(navegador, 20).until(
    EC.visibility_of_element_located(
        (By.XPATH, "//span[text()=' Todos ']")
    )
)
print("clicando status...")
status_forma.click()

# Localiza o botão "Buscar" e clica nele
botao_buscar = WebDriverWait(navegador, 20).until(
    EC.element_to_be_clickable(
        (By.XPATH, '/html/body/app-root/app-home-v2/app-side-bar/mat-sidenav-container/mat-sidenav-content/mat-card/app-saque-esteira/mat-card-content/app-form-saque-consulta-esteira/div/form/app-botao-icon-v2/button')
    )
)
print("apertando botao buscar...")
botao_buscar.click()

# Aguarda por um elemento que indique ausência de resultados
try:
    mensagem_sem_resultado = WebDriverWait(navegador, 40).until(
        EC.presence_of_element_located(
            (By.XPATH, "//span[contains(text(), 'Nenhum arquivo encontrado')]")
        )
    )
    print("Nenhum arquivo encontrado. Encerrando a aplicação.")
    navegador.quit()
except TimeoutException:
    print("Arquivos encontrados ou a mensagem de erro não apareceu. Continuando o script.")


tentativas = 3

for tentativa in range(tentativas):
    try:
        botao_exportar = WebDriverWait(navegador, 20).until(
            EC.element_to_be_clickable(
                (By.XPATH, '/html/body/app-root/app-home-v2/app-side-bar/mat-sidenav-container/mat-sidenav-content/mat-card/app-saque-esteira/mat-card-content/app-form-saque-consulta-esteira/div/app-botao-icon-v2/button')
            )
        )
        print("Botão Exportar clicado!")
        botao_exportar.click()
        break
    except TimeoutException:
        print(f"Tentativa {tentativa +
              1} para clicar no botão Exportar... continuando.")
        if tentativa == tentativas - 1:
            print(
                "Botão Exportar não encontrado após 3 tentativas. Encerrando o navegador.")
            navegador.quit()


##########################################
# Clicar no item menu_tetris
menu_tetris = WebDriverWait(navegador, 10).until(
    EC.element_to_be_clickable(
        (By.XPATH, "//i[contains(@class, 'icon-element-3')]"))
)
print("clicando no tetris...")
menu_tetris.click()


time.sleep(2)
modal_esperado = WebDriverWait(navegador, 10).until(
    # Verifica se o elemento do modal está visível
    EC.visibility_of_element_located(
        (By.XPATH, '//*[@id="dorp_list_icons_0"]'))
)
print("Esperando modal...")
# Agora, clicar no elemento dentro do modal
relatorio_menu = WebDriverWait(navegador, 10).until(
    EC.presence_of_element_located(
        (By.XPATH,
         "//*[@id='dorp_list_icons_1']//mat-icon[@data-mat-icon-name='task-square-bm']"))
)

print("clicando em vendas consignadas...")
relatorio_menu.click()

time.sleep(2)
menu_hamburguer = WebDriverWait(navegador, 10).until(
    EC.element_to_be_clickable(
        (By.XPATH, '/html/body/app-root/app-home-v2/app-master-header-layout/mat-toolbar/button/span[1]/mat-icon'))
)
print("Apertando menu hamburguer...")
menu_hamburguer.click()

relatorio_menu_hamburguer = WebDriverWait(navegador, 10).until(
    EC.presence_of_element_located(
        (By.XPATH, "//mat-panel-title[contains(text(), 'Relatórios')]"))
)
print("Clicando relatorio dentro do menu...")
relatorio_menu_hamburguer.click()


botao_relatorio_esteira = WebDriverWait(navegador, 10).until(
    EC.element_to_be_clickable(
        (By.XPATH,
         "//mat-list-item//a[contains(@href, '/relatorio/esteira-saque')]"))
)
print("Apertando botão 'Relatório Esteira Saque'...")
botao_relatorio_esteira.click()

##############################

try:
    # Espera pelo botão de download e clica
    botao_relatorio_download = WebDriverWait(navegador, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH, "//button[.//mat-icon[text()='download']]")
        )
    )
    print("Apertando botão 'download'...")
    botao_relatorio_download.click()

    # Aguarda um tempo para o download ser concluído
    print("Aguardando o download ser concluído...")
    # Tempo para garantir que o download ocorra (ajuste conforme necessário)
    time.sleep(5)

except TimeoutException:
    print("Botão 'download' não encontrado. Fechando o navegador.")
    navegador.quit()  # Fecha o navegador se o botão não for encontrado
    exit()  # Encerra o script
print("Download (ou tentativa de download) concluído. Fechando a automação.")

navegador.quit()  # Fecha o navegador ao final
