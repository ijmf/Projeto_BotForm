"""
WARNING:

Please make sure you install the bot with `pip install -e .` in order to get all the dependencies
on your Python environment.

Also, if you are using PyCharm or another IDE, make sure that you use the SAME Python interpreter
as your IDE.

If you get an error like:
```
ModuleNotFoundError: No module named 'botcity'
```

This means that you are likely using a different Python interpreter than the one used to install the bot.
To fix this, you can either:
- Use the same interpreter as your IDE and install your bot with `pip install --upgrade -r requirements.txt`
- Use the same interpreter as the one used to install the bot (`pip install --upgrade -r requirements.txt`)

Please refer to the documentation for more information at https://documentation.botcity.dev/
"""


# Import for the Web Bot
from aifc import Error
import shutil
from botcity.web import WebBot, Browser, By

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *
from botcity.web.util import element_as_select
from botcity.web.parsers import table_to_dict
from botcity.plugins.excel import BotExcelPlugin
from botcity.plugins.email import BotEmailPlugin
from pandas import *
import pandas

# Instanciar o plug -in
email = BotEmailPlugin()

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False

excel = BotExcelPlugin()
# excel.add_row(["Nome", "Último", "Máxima", "Mínima",
#               "Variação", "Var. %", "Vol.", "Hora"])


def main():
    # Runner passes the server url, the id of the task being executed,
    # the access token and the parameters that this task receives (when applicable).
    maestro = BotMaestroSDK.from_sys_args()
    # Fetch the BotExecution with details from the task, including parameters
    execution = maestro.get_execution()

    # Se executar pelo VScode comentar o trecho abaixo, executando pelo maestro necessário descomentar.

    # maestro.login(server="https://developers.botcity.dev",
    #               login="57444048-4a34-432e-985f-88d6252065f1",
    #               key="574_SFXHGJ4TTVUWBDXFN6ES")

    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")

    # Obtendo credenciais do Maestro
    # usuario = maestro.get_credential("dados-login", "usuario")
    # senha = maestro.get_credential("dados-login", "senha")

    # Enviando alerta para o Maestro
    # maestro.alert(
    #     task_id=execution.task_id,
    #     title="Iniciando processo",
    #     message=f"O processo de consulta foi iniciado",
    #     alert_type=AlertType.INFO
    # )

    bot = WebBot()

    # Configure whether or not to run on headless mode
    bot.headless = False

    # Uncomment to change the default Browser to Chrome
    bot.browser = Browser.CHROME

    # Uncomment to set the WebDriver path
    bot.driver_path = r"C:\Treinamento BotCity\chromedriver-win64\chromedriver.exe"

    data = pandas.read_excel(
        "C:\Treinamento BotCity\Projetos\EmployeesFeedback.xlsx")
    print(data)

    # Abrimos o Formulário.
    bot.browse("https://forms.gle/tgMfsjvQzwpmU52s9")

    bot.wait(2000)

    for index, row in data.iterrows():

        #Preenchimento do primeiro campo (Texto)
        employee_name_field = bot.find_element(
            r"//*[@id='mG61Hd']/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div/div[1]/input", By.XPATH)
        employee_name_field.send_keys(row['Employee Name'])

        bot.wait(100)

        #Preenchimento do segundo campo (Texto)
        years_of_service_field = bot.find_element(
            r"//*[@id='mG61Hd']/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div/div[1]/input", By.XPATH)
        years_of_service_field.send_keys(row['Years of Service'])

        bot.wait(1000)

        #Preenchimento do Terceiro campo (List)
        departament_field = bot.find_element(
            r"//div[contains(@data-params, 'Department')]//div[contains(@role, 'listbox')]", By.XPATH)
        departament_field.click()

        bot.wait(500)

        #Seleção do Terceiro campo (List)
        departament_field_option = bot.find_element(
            f"//div[@role='option' and @data-value='{row['Department']}']", By.XPATH)
        departament_field_option.click()

        bot.wait(500)

        #Seleção do quarto campo (Radium button)
        employee_satisfaction_field = bot.find_element(
            f"//div[contains(@data-params, 'Employee satisfaction')]//span[text()='{row['Satisfaction Rating']}']", By.XPATH)
        employee_satisfaction_field.click()
        bot.wait(500)
        
        #Clique no botão enviar (Button)
        submit_btn = bot.find_element("//*[@id='mG61Hd']/div[2]/div/div[3]/div[1]/div[1]/div/span/span", By.XPATH)
        submit_btn.click()
        bot.wait(500)
        
        #Clique na opção de responder novamente (Texto com link)
        submit_btn_another = bot.find_element("/html/body/div[1]/div[2]/div[1]/div/div[4]/a[2]", By.XPATH)
        submit_btn_another.click()
        bot.wait(500)
        
        # maestro.new_log_entry(
        #     activity_label="ATIVOS",
        #     values={"NOME": f"{var_Ativo}",
        #             "ULTIMO": f"{var_Ultimo}",
        #             "MAXIMA": f"{var_Maxima}",
        #             "MINIMA": f"{var_Minima}",
        #             "VARIACAO": f"{var_Tot}",
        #             "VAR_POR": f"{var_Por}",
        #             "VOLUME": f"{var_Volume}",
        #             "HORA": f"{var_Hora}",
        #             })

    # Adiciona a linha ao Excel
        # excel.add_row([var_Ativo, var_Ultimo, var_Maxima,
        #               var_Minima, var_Tot, var_Por, var_Volume, var_Hora])

    # excel.write(
    #     r"C:\Treinamento BotCity\Projetos\RelatorioAtivos\Infos_Ativos.xlsx")

    # Configure IMAP com o servidor Hotmail
    # try:
    #     email.configure_imap("outlook.office365.com", 993)

    # # Configure SMTP com o servidor Hotmail
    #     # smtp.office365.com ou smtp-mail.outlook.com
    #     email.configure_smtp("smtp-mail.outlook.com", 587)

    # # Faça login com uma conta de email válida
    #     email.login("junio.str@hotmail.com", "Teste123@")

    # except Exception as e:
    #     print(f"Erro durante a configuração do e-mail: {e}")

    # # Definindo os atributos que comporão a mensagem
    # para = ["junio.str@hotmail.com"]
    # assunto = "Planilha Ativos"
    # corpo_email = ""
    # arquivos = [
    #     r"C:\Treinamento BotCity\Projetos\RelatorioAtivos\Infos_Ativos.xlsx"]

    # # Enviando a mensagem de e -mail
    # email.send_message(assunto, corpo_email, para,
    #                    attachments=arquivos, use_html=True)

    # # Feche a conexão com os servidores IMAP e SMTP
    # email.disconnect()
    # print("Email enviado com sucesso")

    # Subindo arquivo de resultados
    # caminho_arquivo_xlsx = r"C:\Treinamento BotCity\Projetos\RelatorioAtivos\Infos_Ativos.xlsx"
    # caminho_pasta_xlsx = r"C:\Treinamento BotCity\Projetos\RelatorioAtivos"
    # shutil.make_archive(caminho_arquivo_xlsx, 'zip', caminho_pasta_xlsx)
    # maestro.post_artifact(
    #     task_id=execution.task_id,
    #     artifact_name="Infos_Ativos",
    #     filepath=caminho_arquivo_xlsx + ".zip"
    # )

    # Alerta de email
    # maestro.alert(
    #     task_id=execution.task_id,
    #     title="E-mail OK",
    #     message=f"E-mail enviado com sucesso",
    #     alert_type=AlertType.INFO
    # )

    # Implement here your logic...
    ...

    # Wait 3 seconds before closing
    bot.wait(3000)

    # Finish and clean up the Web Browser
    # You MUST invoke the stop_browser to avoid
    # leaving instances of the webdriver open
    bot.stop_browser()

    # Reportando erro ao Maestro
    # maestro.error(
    # task_id=execution.task_id,
    # exception=erro,
    # screenshot="erro.png"
    # )

    # Uncomment to mark this task as finished on BotMaestro
    # maestro.finish_task(
    #     task_id=execution.task_id,
    #     status=AutomationTaskFinishStatus.SUCCESS,
    #     message="Task Finalizada"
    # )


def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()
