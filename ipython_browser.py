#coding: utf-8

# Imports {{{1
import os
from os import path
from selenium import webdriver


# Global variables {{{1
SITE = 'http://ticket.tzar.ru'
login = "david@rusnavigator.ru"
password = "HtjdynUN"


#def pre_captcha(driver) {{{1
def pre_captcha(driver):
    """Действия, которые надо сделать перед вводом капчи"""
    script = """
    var input = document.getElementById("submitBtn");
    var td = input.parentNode;
    var p = document.createElement("p");
    p.setAttribute("id", "captcha_hint");
    p.style.color = "#ff2400";
    p.style['font-size'] = "1em";
    p.style['font-weight'] = "bold";
    p.style.float = "left";
    p.innerHTML = "Введите капчу САМИ!";
    td.insertBefore(p, input);
    """
    driver.find_element_by_id("password").send_keys(password)
    driver.execute_script(script)
    driver.find_element_by_id("recaptcha_response_field").click() # фокус в поле ввода капчи


# Настройка драйвера {{{1
chromedriver = path.expanduser('~/bin/chromedriver')
os.environ['webdriver.chrome.driver'] = chromedriver
CO = webdriver.ChromeOptions()
CO.binary_location = '/opt/google/chrome/google-chrome'

driver = webdriver.Chrome(chromedriver, chrome_options=CO)


# Открытие браузера -- страница авторизации {{{1
driver.get(SITE)

driver.find_element_by_id("login").send_keys(login)
# Ждём ввода капчи и загрузки главной страницы
pre_captcha(driver)

