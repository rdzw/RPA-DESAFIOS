from botcity.web import WebBot, Browser, By
from seleniumbase import Driver
from time import sleep
import pandas as pd

def main():

    # Inicializa o WebBot (se você quiser usar isso para algo específico mais tarde)
    bot = WebBot()
    bot.headless = False  # Defina como True se preferir rodar em modo headless
    bot.driver_path = 'resources/chromedriver.exe'
    sleep(3)

    # Cria uma instância do Driver (seleniumbase) com undetected_chromedriver (uc)
    driver = Driver(uc=True, headless=False)

    # Navega para a URL usando o seleniumbase
    driver.get("https://www.lme.com/Account/Login")
    sleep(10)

    # Maximiza a janela do navegador
    driver.maximize_window()

    # Aguardar até que o campo de usuário esteja disponível
    try:
        login = driver.find_element(By.XPATH, '/html/body/div[2]/div[3]/div/div[1]/div/div[2]/div/button[2]')
        login.click()
        sleep(2)
        print("Cookie clicado.")
    except Exception as e:
        print(f"Erro: {e}")
        sleep(2)

    while True:
        try:
            email_field = driver.find_element(By.CSS_SELECTOR, '#Email')
            email_field.click()
            sleep(2)
            print('Campo encontrado')
            email_field.send_keys('ademar.castro.fh@gmail.com')
            break
        except Exception as e:
            print(f'Campo de usuário não encontrado, tentando novamente: {e}')
            sleep(1)

    # Aguardar até que o campo de senha esteja disponível
    while True:
        try:
            password_field = driver.find_element(By.XPATH, '/html/body/main/div/div[1]/div/form/div[2]/input')
            password_field.click()
            sleep(2)
            password_field.send_keys('qyiaXnRN9EfC3J!')
            break
        except Exception as e:
            print(f"Campo de senha não encontrado, tentando novamente: {e}")
            sleep(1)

    try:
        driver.find_element(By.CSS_SELECTOR, 'body > main > div > div.form > div > form > div.form__block > button').click()
        sleep(5)
    except Exception as e:
        print(f'Campo de login não encontrado, tentando novamente: {e}')
        sleep(1)

    # Continue com as interações usando o WebBot
    try:
        data = driver.find_element("/html/body/header/div[1]/div/div/div[2]/nav/ul/li[4]/button", By.XPATH)
        data.click()
    except Exception as e:
        print(f"Erro ao clicar no elemento Data: {e}")
    sleep(3)

    try:
        report = driver.find_element("/html/body/header/div[1]/div/div/div[2]/nav/ul/li[4]/div/ul/li[3]/button/span", By.XPATH)
        report.click()
    except Exception as e:
        print(f"Erro ao clicar no elemento Report: {e}")
    sleep(3)

    try:
        monthly = driver.find_element("/html/body/header/div[1]/div/div/div[2]/nav/ul/li[4]/div/ul/li[3]/div/div[2]/ul/li[2]/a", By.XPATH)
        monthly.click()
    except Exception as e:
        print(f"Erro ao clicar no elemento Monthly: {e}")
    sleep(3)

    try:
        primeiro_arquivo = driver.find_element("/html/body/main/div/div[1]/div[1]/div[2]/div/p[1]/a", By.XPATH)
        primeiro_arquivo.click()
        sleep(2)
    except Exception as e:
        print(f"Erro ao tentar baixar o primeiro arquivo: {e}")
        sleep(2)
        
    # Caminho para o arquivo Excel
    caminho_arquivo_excel = r"downloaded_files\July 2024 No Steel  Molybdenum.xlsx"

    ## Obter todos os nomes das planilhas
    planilhas = pd.ExcelFile(caminho_arquivo_excel).sheet_names
    print(planilhas)
    
    # Ler o arquivo Excel (especificando a planilha, se houver mais de uma)
    df = pd.read_excel(caminho_arquivo_excel, sheet_name="ABR") 
    # Exibir os primeiros registros do DataFrame
    print(df.head())
    
    # Exibir os nomes das colunas
    print("Colunas disponíveis:", df.columns)
    
    # Acessar as colunas específicas
    colunas_dados = df[["Unnamed: 1", "Unnamed: 2", "Unnamed: 3", "Unnamed: 4"]]

    # Exibir os dados da coluna
    print(colunas_dados)
    
    
    # try:
    #     segundo = driver.find_element("/html/body/main/div/div[1]/div[1]/div[2]/div/p[1]/a", By.XPATH)
    #     segundo.click()
    #     sleep(2)
    # except Exception as e:
    #     print(f"Erro ao tentar baixar o segundo arquivo: {e}")
    #     sleep(2)

    # try:
    #     terceiro = driver.find_element("/html/body/main/div/div[1]/div[1]/div[2]/div/p[1]/a", By.XPATH)
    #     terceiro.click()
    #     sleep(2)
    # except Exception as e:
    #     print(f"Erro ao tentar baixar o terceiro arquivo: {e}")
    #     sleep(2)

    # Finaliza e limpa os recursos
    driver.quit()
    bot.stop_browser()

if __name__ == '__main__':
    main()
