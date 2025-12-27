import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoAlertPresentException


url = "https://servicioenlinea.ug.edu.ec/SIUG/Account/Login.aspx"

driver = webdriver.Chrome()
driver.get(url)
assert "SIUG" in driver.title
ventana_siug = driver.current_window_handle

try:
    # Busqueda de elementos HTMLL
    i_usuario = driver.find_element(By.NAME, "ctl00$MainContent$LoginUser$UserName")
    i_password = driver.find_element(By.NAME, "ctl00$MainContent$LoginUser$Password")
    i_dia = driver.find_element(By.NAME, "ctl00$MainContent$LoginUser$TXT_DIA")
    i_mes = driver.find_element(By.NAME, "ctl00$MainContent$LoginUser$TXT_MES")
    btn_login = driver.find_element(By.NAME, "ctl00$MainContent$LoginUser$LoginButton")
    
    # Asignacion de valores
    i_usuario.send_keys("1311893711")
    i_password.send_keys("Fer27112002")
    i_dia.send_keys("27")
    i_mes.send_keys("11")
    btn_login.click()

    time.sleep(3)
    
    btn_close = driver.find_element(By.CLASS_NAME, "close")
    btn_close.click()
    
    btn_horario = driver.find_element(By.NAME, "ctl00$MainContent$BTN_CONSULTA_HORARIO")
    btn_horario.click()

    time.sleep(3)
    select_carrera = Select(driver.find_element(By.NAME, "ctl00$MainContent$DDL_CARRERA"))

    print('Entrando al select de carrera')
    select_carrera.select_by_value("0311  ")
    
    time.sleep(3)

    select_periodo = Select(driver.find_element(By.NAME, "ctl00$MainContent$DDL_PLECTIVO"))
    print('Seleccionando el periodo')
    select_periodo.select_by_value('35')

    time.sleep(3)

    btn_clases = driver.find_element(By.ID, 'MainContent_BTN_CLASES')
    print("Haciendo click en el boton de 'Clases Virtuales'")
    btn_clases.click()

    time.sleep(2)

    for handle in driver.window_handles:
        if handle != ventana_siug:
            driver.switch_to.window(handle)
            break

    time.sleep(5)

    btn_ingreso_clase = driver.find_element(By.ID, 'btnIrReunion')
    btn_ingreso_clase.click()
    
    time.sleep(5)

    alert = driver.switch_to.alert
    alert.accept()

    time.sleep(50)


except Exception as e:
    print(f"Ocurrió un error: {e}")

