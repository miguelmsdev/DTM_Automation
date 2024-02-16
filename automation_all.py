# -*- coding: utf-8 -*-

from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import selenium.webdriver.support.expected_conditions as EC
from selenium.common.exceptions import (
NoSuchAttributeException,
NoSuchElementException,
ElementNotInteractableException,
ElementNotVisibleException,
StaleElementReferenceException,
ElementClickInterceptedException,
NoSuchCookieException,
)
from decimal import Decimal, ROUND_DOWN
from datetime import datetime
import pandas as pd
import traceback
import re
import random
import time
import sys
import subprocess

from pynput import keyboard

from personal_data.config import CHROME_PROFILE_PATH, PASSWORD
from IPython.display import display

def run_automation(company_name, relatorio_path, scans_path, start_date_value, end_date_value):
    options = Options()
    options.add_argument(CHROME_PROFILE_PATH)
    options.add_argument("--lang=en")
    options.add_argument("--debuggerAddress=localhost:9348")
    options.add_argument("--remote-debugging-port=9348")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--disable-logging")
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("detach", True)

    service = ChromeService(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    wait = WebDriverWait(driver, timeout=50, poll_frequency=0.1)

    # ------------------------------global Functions -----------------------------



    def loginMTR():
        driver.find_element(By.ID, "mat-input-0").send_keys("04737150000133")
        time.sleep(random.uniform(0.1, 0.2))
        click(driver, (By.ID, "mat-input-1"))
        time.sleep(random.uniform(1, 1.5))
        send_Keys(driver, (By.ID, "mat-input-1"), ("56073623704"))
        time.sleep(random.uniform(0.1, 0.2))
        driver.find_element(By.ID, "mat-input-2").send_keys(PASSWORD)
        time.sleep(random.uniform(1, 1.5))
        driver.find_element(
            By.XPATH,
            "/html/body/app-root/app-inicio/mat-sidenav-container/mat-sidenav-content/mat-card/mat-card-content/div[2]/div/mat-card/mat-card-content/form/div[2]/div/button",
        ).click()


    def searchMTRInfo():
        time.sleep(2)
        print(start_date_value)
        send_Keys(driver, (By.CSS_SELECTOR, 'input[formcontrolname="manDataInicialDestinador"]'), start_date_value)
        print(end_date_value)
        send_Keys(driver, (By.CSS_SELECTOR, 'input[formcontrolname="manDataFinalDestinador"]'), end_date_value)

        send_Keys(driver, (By.CSS_SELECTOR, 'input[formcontrolname="buscaDestinador"]'), company_cnpj)
        click(driver, (By.XPATH,
            "/html/body/app-root/app-navegacao/mat-sidenav-container/mat-sidenav-content/app-meus-mtrs/mat-sidenav-container/mat-sidenav-content/mat-card[2]/form/div[3]/button")) # Pesquisar Button
        time.sleep(1)
        # Wait for page 2 button to appear, means page is loaded
        while (
            len(
                driver.find_elements(
                    By.XPATH,
                    "/html/body/app-root/app-navegacao/mat-sidenav-container/mat-sidenav-content/app-meus-mtrs/mat-sidenav-container/mat-sidenav-content/mat-card[2]/form/p-table/div/p-paginator/div/span/a[2]",
                )
            )
            < 1
        ):
            time.sleep(random.uniform(0.05, 0.1))


    def to_comma_to_4_digit_after(to_be_converted: str) -> str:
        number = Decimal(to_be_converted.replace(",", "."))
        formatted_number = "{:.4f}".format(number.quantize(Decimal('0.0000'), rounding=ROUND_DOWN)).replace(".", ",")
        return formatted_number


    def goToFinalPage():
        button_2_page = driver.find_element(
            By.XPATH,
            "/html/body/app-root/app-navegacao/mat-sidenav-container/mat-sidenav-content/app-meus-mtrs/mat-sidenav-container/mat-sidenav-content/mat-card[2]/form/p-table/div/p-paginator/div/a[4]",
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", button_2_page)
        button_2_page.click()  # Página final


    def getUiStateActiveValue():
        time.sleep(random.uniform(0.2, 0.5))
        while (
            len(
                driver.find_elements(
                    By.XPATH,
                    f"/html/body/app-root/app-navegacao/mat-sidenav-container/mat-sidenav-content/app-meus-mtrs/mat-sidenav-container/mat-sidenav-content/mat-card[2]/form/p-table/div/p-paginator/div/span/a[2]",
                )
            )
            < 1
        ):
            time.sleep(random.uniform(0.05, 0.1))

        while len(driver.find_elements(By.CLASS_NAME, "ui-state-active")) < 1:
            time.sleep(random.uniform(0.05, 0.1))

        ui_state_active_element = driver.find_element(
            By.CSS_SELECTOR, ".ui-state-active.ui-paginator-page"
        )
        ui_state_active_value = int(ui_state_active_element.get_attribute("text"))
        # isPageHigherLowerSame(ui_state_active_value)
        return ui_state_active_value


    def changingPageDown():
        # time.sleep(1)
        button_page_down = driver.find_element(
            By.XPATH,
            "/html/body/app-root/app-navegacao/mat-sidenav-container/mat-sidenav-content/app-meus-mtrs/mat-sidenav-container/mat-sidenav-content/mat-card[2]/form/p-table/div/p-paginator/div/a[2]",
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", button_page_down)
        button_page_down.click()


    def receivingMTR():
        emitted_date_element = driver.find_element(
            By.XPATH,
            "/html/body/app-root/app-navegacao/mat-sidenav-container/mat-sidenav-content/app-meus-mtrs/mat-sidenav-container/mat-sidenav-content/mat-card[2]/form/p-table/div/div/table/tbody/tr[1]/td[2]",
        )
        unprocessed_received_date = emitted_date_element.get_attribute("innerHTML")
        almost_processed_received_date = datetime.strptime(
            unprocessed_received_date, "%m/%d/%Y"
        )
        emitted_date = almost_processed_received_date.strftime("%d/%m/%Y")
        # print(received_date)

        # print(f"table_received_date's type is: {type(table_received_date)}")
        # print(f"received_date's type is: {type(received_date)}")

        time.sleep(random.uniform(0.1, 0.3))

        mtr_xpath_row_number = 0

        for i in range(10):
            mtr_xpath_row_number += 1

            mtr_number_element = getText(driver, (By.XPATH, f"/html/body/app-root/app-navegacao/mat-sidenav-container/mat-sidenav-content/app-meus-mtrs/mat-sidenav-container/mat-sidenav-content/mat-card[2]/form/p-table/div/div/table/tbody/tr[{mtr_xpath_row_number}]/td[1]"))

            time.sleep(random.uniform(0.1, 0.3))

            unprocessed_mtr_number = mtr_number_element
            # unprocessed_mtr_number = mtr_number_element.get_attribute("innerHTML")
            mtr_number = re.sub(r"[^0-9]+", "", unprocessed_mtr_number)

            
            mtr_situation = getText(driver, (By.XPATH, f"/html/body/app-root/app-navegacao/mat-sidenav-container/mat-sidenav-content/app-meus-mtrs/mat-sidenav-container/mat-sidenav-content/mat-card[2]/form/p-table/div/div/table/tbody/tr[{mtr_xpath_row_number}]/td[5]"))
            time.sleep(random.uniform(0.1, 0.3))
            if mtr_situation == "Salvo":
                time.sleep(random.uniform(0.1, 0.3))
                print(f"Row: {mtr_xpath_row_number} - Nº {mtr_number} - \N{CHECK MARK} IS in mtr_scans_dict")
                # random driver and random plate
                random_driver, random_plate = get_random_values(drivers_dict_dict)

                click(driver, (By.XPATH, f"/html/body/app-root/app-navegacao/mat-sidenav-container/mat-sidenav-content/app-meus-mtrs/mat-sidenav-container/mat-sidenav-content/mat-card[2]/form/p-table/div/div/table/tbody/tr[{mtr_xpath_row_number}]/td[6]/a[1]"))

                # procurando por responsavel aloisio visivel (se já foi escolhido na sessão)
                if (
                    len(
                        driver.find_elements(
                            By.XPATH,
                            "/html/body/app-root/app-navegacao/mat-sidenav-container/mat-sidenav-content/app-meus-mtrs/mat-sidenav-container/mat-sidenav-content/p-dialog[2]/div/div[2]/form/mat-card/div[2]/div[2]",
                        )
                    )
                    < 1
                ):
                    time.sleep(random.uniform(0.1, 0.2))

                    # selecionar responsavel
                    click(driver, (By.XPATH, "/html/body/app-root/app-navegacao/mat-sidenav-container/mat-sidenav-content/app-meus-mtrs/mat-sidenav-container/mat-sidenav-content/p-dialog[2]/div/div[2]/form/mat-card/div[2]/div/button"))
                    # simbolo de confirmar ação de cargo
                    click(driver, (By.XPATH, "/html/body/app-root/app-navegacao/mat-sidenav-container/mat-sidenav-content/app-meus-mtrs/mat-sidenav-container/mat-sidenav-content/p-dialog[7]/div/div[2]/p-table/div/div[2]/table/tbody/tr/td[3]/a[2]"))

                # Data de recebimento

                click(driver, (By.CSS_SELECTOR, "input[placeholder='Data de Recebimento']"))
                send_Keys(driver, (By.CSS_SELECTOR, "input[placeholder='Data de Recebimento']"), select_delete)
                send_Keys(driver, (By.CSS_SELECTOR, "input[placeholder='Data de Recebimento']"), emitted_date)

                # Data de recebimento

                # Motorista
                print(          f"Random Driver: {random_driver} - Random Plate: {random_plate}")

                driver_element = driver.find_element(
                    By.CSS_SELECTOR, "input[placeholder='Motorista']"
                )# Motorista

                print(driver_element.text)
                # driver_element_text = driver_element.text

                if driver_element.text in ["", None, " "] or not isinstance(driver_element.text, str):
                    send_Keys(driver, (By.CSS_SELECTOR, "input[placeholder='Motorista']"), random_driver)

                plate_element = driver.find_element(
                    By.CSS_SELECTOR, "input[placeholder='Placa']"
                ) # Placa

                print(plate_element.text)
                # plate_element_text = plate_element.text
                
                if plate_element.text in ["", None, " "] or not isinstance(plate_element.text, str):
                    send_Keys(driver, (By.CSS_SELECTOR, "input[placeholder='Placa']"), random_plate)
                    

                time.sleep(random.uniform(0.2, 0.5))

                tons_to_receive = mtr_dict[mtr_number]["table_tons"]

                driver.find_element(
                    By.CSS_SELECTOR, "input[formcontrolname='marQuantidadeRecebida']"
                ).send_keys(
                    tons_to_receive
                )  # Toneladas recebidas

                time.sleep(random.uniform(0.1, 0.3))

                # Receber Finalizar
                driver.find_element(
                    By.XPATH,
                    f"/html/body/app-root/app-navegacao/mat-sidenav-container/mat-sidenav-content/app-meus-mtrs/mat-sidenav-container/mat-sidenav-content/p-dialog[2]/div/div[3]/p-footer/div/div[1]/button",
                ).click()



                click(driver, (By.CSS_SELECTOR, "div.close-button.ng-star-inserted"))


                print(          f"MTR Received: {tons_to_receive} tons")

                time.sleep(random.uniform(0.5, 0.7))

                # time.sleep(0.3)  # unnecessary
            elif mtr_situation == "Recebido":
                continue
        if mtr_xpath_row_number == 10 and ui_state_active_value == 1:
            print("page 1, last mtr done,quitting in 30 seconds")
            time.sleep(30)
            driver.quit()
        else:
            changingPageDown()


    def change_date_format(date_string):
        if not isinstance(date_string, str):
            return ""  # or any other value you prefer for non-date cells
        try:
            # Parse the input string into a datetime object
            datetime_obj = datetime.strptime(date_string, "%Y-%m-%d %H:%M:%S")

            # Format the datetime object into the desired format
            new_date_format = datetime_obj.strftime("%d/%m/%Y")

            return new_date_format
        except ValueError:
            # print(ValueError)
            return ""  # or any other value you prefer for non-date cells


    def get_random_values(dictionary):
        random_driver = random.choice(list(dictionary[company_name].keys()))
        while random_driver == "":
            random_driver = random.choice(list(dictionary[company_name].keys()))
        random_plate = dictionary[company_name][random_driver]
        return random_driver, random_plate


    def click(driver, locator):
        # driver.execute_script("arguments[0].scrollIntoView(true);", element)
        # element.click()
        max_retries = 100
        retry_count = 0
        while retry_count < max_retries:
            try:
                driver.execute_script("""
                arguments[0].scrollIntoView(true);
                arguments[0].click();
                """, wait.until(EC.presence_of_element_located(locator)))
                break  # Break out of the while loop if the code execution is successful
            except Exception as e:
                # Print the error message
                print(f"Error: {str(e)}retrying...")
                time.sleep(0.20)
                # Increment the retry count
                retry_count += 1

    def send_Keys(driver, locator, value):
        element = wait.until(EC.presence_of_element_located(locator))
        element.send_keys(select_delete)
        # element.clear()
        element.send_keys(value)
        # driver.execute_script("arguments[0].scrollIntoView(true);", element)
        # element.send_keys(value)


    def getText(driver, locator):
        max_retries = 100
        retry_count = 0
        while retry_count < max_retries:
            try:
                element = wait.until(EC.presence_of_element_located(locator))
                driver.execute_script("""
                arguments[0].scrollIntoView(true);
                arguments[0].click();
                """, element)

                return element.text  # Return the text of the element

            except Exception as e:
                # Print the error message
                print(f"Error: {str(e)}, retrying...")
                time.sleep(0.20)
                # Increment the retry count
                retry_count += 1

        return None  # Return None if the element is not found after retries

    def convert_path(path):
        converted_path = path.replace("/", r"\\")
        return converted_path


    stop_flag = False

    def on_press(key):
        global stop_flag
        if key == keyboard.Key.f10:
            stop_flag = True
            return False
        


    listener = keyboard.Listener(on_press=on_press)
    listener.start()


    # ------------------------------global Functions -----------------------------
    # ------------------------------reading Table 1 --------------------------------
    print(f"{convert_path(relatorio_path)}")
    tabela = pd.read_excel(convert_path(relatorio_path))
    # display(tabela)

    mtr_dict = {}
    mtr_dict.clear()


    for linha in tabela.index:
        # pegar nº de MTR como chave e toneladas indicadas e cnpj do gerador 
        table_mtr_number = str(tabela.loc[linha, "Nº MTR"])
        unprocessed_table_tons = str(tabela.loc[linha, "Quantidade indicada"])
        table_tons = to_comma_to_4_digit_after(unprocessed_table_tons)
        table_cnpj = str(tabela.loc[linha, "Gerador (CNPJ/CPF)"])


        mtr_table_data = {"table_tons": table_tons, "table_cnpj": table_cnpj}
        mtr_dict[table_mtr_number] = mtr_table_data

    # ------------------------------reading Table 1 --------------------------------

    # ------------------------------reading Table 2 --------------------------------
    tabela = pd.read_excel(
        convert_path(scans_path), sheet_name="Scans"
    )

    mtr_scans_dict = {}
    mtr_scans_dict.clear()

    for linha in tabela.index:
        # pegar dados
        table_mtr_scan = str(tabela.iloc[linha, 2])
        if len(table_mtr_scan) <= 4:
            continue
        unformatted_table_received_date = str(tabela.iloc[linha, 1])

        table_received_date = change_date_format(unformatted_table_received_date)
        # print(table_mtr_scan)
        # print(table_received_date)
        mtr_scans_dict[table_mtr_scan] = table_received_date

    print(mtr_scans_dict)

    # ------------------------------reading Table 2 --------------------------------

    # -------------------------------- send_keys ---------------------------------

    select_delete = Keys.CONTROL + "a", Keys.DELETE
    select = Keys.CONTROL + "a"
    select_copy = Keys.CONTROL + "a", Keys.CONTROL + "c"
    copy = Keys.CONTROL + "c"
    paste = Keys.CONTROL + "v"

    # -------------------------------- send_keys ---------------------------------

    # ------------------------------global Variables -----------------------------

    traceback_str1 = ""
    traceback_str2 = ""


    cnpj_dict = {
        "Telar": "62570320000134",
        "Comér": "27170703000114",
        "Cinco Estrelas": "30686869000100",
        "Consórcio DBO": "41.018.034/0001-90"
        }

    start_date = start_date_value
    end_date = end_date_value
    company_cnpj = cnpj_dict[company_name]

    mtr_table_data = {"table_tons": table_tons, "table_cnpj": table_cnpj}
    mtr_dict[table_mtr_number] = mtr_table_data



    telar_drivers_dict = {
        "Jaime": "LPU4J99",
        "Juliane": "LIP1034",
        "Fábio": "MTH8364",
        "": "MTA8366",
        "Angelo": "MPZ3I55",
        "Castro": "JLL5692",
        "": "OYH8806",
        "Helivander": "AAK6A83",
        "": "MTW5D00",
        "Teofilo": "NRF1432"
    }

    comer_drivers_dict = {
        "Fabricio Rafael Matos de Freitas": "MTY1H29",
        "Paulo Henrique Silva de Araújo": "HFD5H65",
        "Jocielio de Jesus Novaes": "HFD5H66",
        "": "",
        "": "",
        "": "",
        "": "",
        "": "",
        "": ""
    }

    cinco_estrelas_drivers_dict = {
        "Marco": "CNI6A72",
        "Nelson": "MTS3104",
        "Lucas": "NZA7A31",
        "Felipe": "IRY2J55",
        "André": "MSJ1I98",
        "Bruno": "OLQ7A12",
        "Marlon": "MPU8E48",
        "Renilton": "FDC4F90",
        "Breno": "MRC5445",
        "Luan": "NJO5373"
    }

    consorcio_dbo_drivers_dict = {
        "José Mota  ": "CNI6A72",
        "José Jorge": "MTS3104",
        "Erikson Tomaz": "NZA7A31",
        "Luiz Carlos": "IRY2J55",
        "Amarildo": "MSJ1I98"
    }


    drivers_dict_dict = {
        "Telar": telar_drivers_dict,
        "Comér": comer_drivers_dict,
        "Cinco Estrelas": cinco_estrelas_drivers_dict,
        "Consórcio DBO": consorcio_dbo_drivers_dict
    }


        # ------------------------------global Variables -----------------------------

    print("Logging in...")

    # getting to page, maximizing window
    driver.get("https://mtr.sinir.gov.br/#/inicio")
    driver.maximize_window()

    time.sleep(1)

    loginMTR()

    time.sleep(2)

    driver.get("https://mtr.sinir.gov.br/#/navegacao/meusmtrs")

    searchMTRInfo()

    goToFinalPage()

    try:
        while not stop_flag:
            ui_state_active_value = getUiStateActiveValue()
            print(f"Active Page: {ui_state_active_value}")

            if len(driver.find_elements(By.CSS_SELECTOR, "[title='Receber MTR']")) >= 1:
                receivingMTR()
            else:
                changingPageDown()
    except (
        Exception,
        NoSuchAttributeException,
        NoSuchElementException,
        ElementNotInteractableException,
        ElementNotVisibleException,
        StaleElementReferenceException,
        ElementClickInterceptedException,
        NoSuchCookieException,
        ):
        traceback_str2 = traceback.format_exc()

        with open('traceback_log.txt', 'a') as file:
            current_datetime = datetime.now()
            file.write(f"Date and Time: {current_datetime}\n")
            # if traceback_str1 == "":
            #     file.write("No Traceback 1 exception.\n")
            # else:
            #     file.write("Traceback 1:\n")
            #     file.write(traceback_str1)
            if traceback_str2 == "":
                file.write("\nNo Traceback 2 exception.\n")
            else:
                file.write("\n\nTraceback 2:\n")
                file.write(traceback_str2)
                file.write("\n------------------------------\n\n")

        traceback.print_exc()
        print("an error stopped the code")
        # driver.get("https://www.youtube.com/watch?v=HrGjqPhzErs")
        should_code_end = input("Error occurred, do you want to close the driver? Y/N")
        if should_code_end in ["Y", "Yes", "y", "yes"]:
                print("180 seconds left")
                time.sleep(60)
                print("120 seconds left")
                time.sleep(60)
                print("60 seconds left")
                time.sleep(60)
                should_code_end = input("Final warning, do you want to close the driver? Y/N")
                if should_code_end in ["Y", "Yes", "y", "yes"]:
                    print("180 seconds left")
                    time.sleep(60)
                    print("120 seconds left")
                    time.sleep(60)
                    print("60 seconds left")
                    time.sleep(60)
                    driver.quit()
                else:
                    driver.quit()
        elif should_code_end in ["N", "No", "n", "no"]:
            driver.quit()
        else:
            print("typed wrong, waiting 180 seconds")
            time.sleep(180)



    print("code ended")
    time.sleep(180)
