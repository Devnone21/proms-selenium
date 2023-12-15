import json
import os
import time
import customtkinter as ctk
from tkinter import CENTER
from tkinter import filedialog as fd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from datetime import datetime
import logging

logging.basicConfig(
    filename=f'out.log.{datetime.now().strftime("%m%d%H%M")}', filemode='a', level=logging.DEBUG,
    datefmt='%Y-%m-%d %H:%M:%S',
    format='%(asctime)s, %(levelname)s, %(message)s')

# Load configuration e.g driver path, web info
config = json.load(open('config.json'))

# Basic parameters and initializations
# Supported modes : Light, Dark, System
ctk.set_appearance_mode(config.get('Preferences', {}).get('app_mode', 'System'))
# Supported themes : green, dark-blue, blue
ctk.set_default_color_theme(config.get('Preferences', {}).get('color_theme', 'blue'))
# browser settings and open WEB_URL
options = Options()
options.add_argument('start-maximized')
options.add_argument("--disable-extensions")
options.add_argument('--no-sandbox')
options.add_argument("--disable-gpu")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--dns-prefetch-disable")
options.add_experimental_option('prefs', {
    'credentials_enable_service': False,
    'profile.password_manager_enabled': False,
    'excludeSwitches': ['enable-automation'],
    'useAutomationExtension': False,
})
options.binary_location = config.get("CHROME").get("binary_path")


class BROWSER:
    """ Commonly use chrome webdriver functions """
    def __init__(self):
        self.driver = webdriver.Chrome(
            service=Service(executable_path=config.get("CHROME").get("driver_path")),
            options=options)
        self.driver.set_page_load_timeout(10)

    def browser_xpathclick(self, xpath):
        btn = self.driver.find_element(By.XPATH, xpath)
        self.driver.execute_script("arguments[0].click();", btn)
        WebDriverWait(self.driver, 12).until(lambda d: d.execute_script("return jQuery.active == 0"))
        return btn

    def browser_input(self, xpath, text):
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((
            By.XPATH, xpath))).send_keys(text)

    def browser_scrolldown(self):
        self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")


# ==================== CTk App Utility functions =====================
def clean(lst: list):
    clean_lst = [s.strip(',').strip('"') for s in lst]
    return [x for x in list(set(clean_lst)) if len(x) > 5]


def list_to_entry(list_str: list) -> str:
    tasks = clean(list_str)
    return ',\n'.join(f'"{n}"' for n in tasks)


def entry_to_list(str_entry: str) -> list:
    list_str = clean(str_entry.split('\n'))
    return list_str


def extract_ref_no(inner_text: str):
    split_dash = inner_text.split('-')
    if not split_dash:
        return ''
    split_colon = split_dash[-1].split('/')
    if not split_colon:
        return ''
    return split_colon[-1]


# ========================= CTk App Class ============================
class App(ctk.CTk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.web = None
        self.projects = []
        self.plan_start = '01/09/2023'

        self.title("Auto Create Work Order - Proms")
        self.geometry("600x400")
        self.minsize(600, 400)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure((0, 2), weight=1)
        # File Upload Label
        self.nameLabel = ctk.CTkLabel(self, text="Project File list")
        self.nameLabel.grid(row=0, column=0, columnspan=3, padx=20, pady=(20, 0), sticky="s")
        # File Upload List
        self.filelistBox = ctk.CTkTextbox(self, height=220)
        self.filelistBox.place(relx=0.5, rely=0.5, anchor=CENTER)
        self.filelistBox.grid(row=1, column=0, columnspan=3, padx=10, pady=20, sticky="nsew")
        # Add Files Button
        self.addButton = ctk.CTkButton(self, text="1. Choose file(s)", command=self.click_select_file)
        self.addButton.grid(row=3, column=0, padx=20, pady=(0, 20), sticky="w")
        # Plan Start Label
        self.planLabel = ctk.CTkLabel(self, text="2. Set Plan Start Date: ")
        self.planLabel.grid(row=3, column=1, padx=(150, 0), pady=(0, 20), sticky="ne")
        # Plan Start EntryBox
        self.entryDate = ctk.CTkEntry(self, placeholder_text="01/09/2023")
        self.entryDate.grid(row=3, column=2, padx=20, pady=(0, 20), sticky="e")
        # Tasks Run Button
        self.runButton = ctk.CTkButton(self, text="3. Run", command=self.run_automation)
        self.runButton.grid(row=4, column=2, padx=20, pady=(0, 20), sticky="e")

    def click_select_file(self):
        """ Command when click Choose file(s) button """
        filetypes = (
            ('Project files', '*.xlsx'),
            ('All files', '*.*')
        )
        filenames = fd.askopenfilenames(title='Open files', initialdir='/', filetypes=filetypes)
        current_text = self.filelistBox.get("0.0", "end")
        self.projects = entry_to_list(current_text)
        logging.info(f'on_add => cur: {self.projects}, new: {filenames}')
        self.projects.extend(list(filenames))
        self.filelistBox.delete("0.0", "end")
        self.filelistBox.insert("0.0", list_to_entry(self.projects))

    def auto_create_wo(self, order, project_excel):
        """ Sub-function to create Work Order after Run automation """
        project_id = project_excel.loc[order, 'Project ID'].value
        proms_site = project_excel.loc[order, 'Proms Site'].value
        proms_node = "" # project_excel.loc[order, 'Proms Node'].value
        step_nth = [False] * 4
        try:
            # click button "New Work Order", wait for page load
            self.web.browser_xpathclick('//button[@id="newWorkOrderBtn"]')
            time.sleep(2)
            # select dropdown "Core Online"
            dropdown = self.web.browser_xpathclick('//select[@id="template_sos_type"]')
            option_value = self.web.driver.find_element(
                By.XPATH,
                '//select[@id="template_sos_type"]//option[contains(text(), "Core Online")]'
            ).get_attribute("value")
            Select(dropdown).select_by_value(option_value)
            time.sleep(2)
            # input SiteID (Proms Site), click outside, wait
            self.web.browser_input('//input[@id="siteNodeId"]', proms_site)
            time.sleep(2)
            # input Site From (Proms Node), wait
            # self.web.browser_input('//input[@id="s_from"]', proms_node)
            # time.sleep(2)
            # click dropdown button "projectTypeList", wait
            self.web.browser_xpathclick('//div[@id="template_projectType"]//button')
            time.sleep(1)
            # click option button "ACC MPLS"
            self.web.browser_xpathclick('//div[@id="template_projectType"]//span[contains(text(), "ACC MPLS")]')
            time.sleep(1)
            self.web.browser_xpathclick('//div[@id="template_projectType"]//button')
            # click select dropdown option, title="200070 - Huawei IPRAN DN expansion"
            dropdown = self.web.browser_xpathclick('//select[@id="ddlProject"]')
            option_value = self.web.driver.find_element(
                By.XPATH,
                f'//select[@id="ddlProject"]//option[contains(text(), "{project_id}")]'
            ).get_attribute("value")
            Select(dropdown).select_by_value(option_value)
            time.sleep(2)
            # click select dropdown option, contains "TRUE INTERNET CORPORATION CO."
            dropdown = self.web.browser_xpathclick('//select[@id="ddlCompany"]')
            option_value = self.web.driver.find_element(
                By.XPATH,
                '//select[@id="ddlCompany"]//option[contains(text(), "TRUE MOVE H UNIVERSAL COMMUNICATION")]'
            ).get_attribute("value")
            Select(dropdown).select_by_value(option_value)
            time.sleep(1)
            # click select dropdown option, contains "CO104"
            dropdown = self.web.browser_xpathclick('//select[@id="ddlWoTemplate"]')
            option_value = self.web.driver.find_element(
                By.XPATH,
                '//select[@id="ddlWoTemplate"]//option[contains(text(), "CO604")]'
            ).get_attribute("value")
            Select(dropdown).select_by_value(option_value)
            time.sleep(1)
            # click button "Get WO Template", wait for page load
            self.web.browser_xpathclick('//button[@id="template_btn_template"]')
            time.sleep(2)
            step_nth[0] = True

            # click checkbox, wait for option expand
            self.web.browser_xpathclick('//input[@id="checkbox0"]')
            time.sleep(1)
            # click + button, wait, scroll down
            self.web.browser_xpathclick('//button[@id="btnWoType0"]')
            time.sleep(1)
            self.web.browser_scrolldown()
            # click un-check TK02-TK06
            self.web.browser_xpathclick('//div[@id="demo0"]//td[contains(text(), "NT_IE_TK02")]/..//input')
            self.web.browser_xpathclick('//div[@id="demo0"]//td[contains(text(), "NT_IE_TK03")]/..//input')
            self.web.browser_xpathclick('//div[@id="demo0"]//td[contains(text(), "NT_IE_TK04")]/..//input')
            self.web.browser_xpathclick('//div[@id="demo0"]//td[contains(text(), "NT_IE_TK05")]/..//input')
            self.web.browser_xpathclick('//div[@id="demo0"]//td[contains(text(), "NT_IE_TK06")]/..//input')
            # select dropdown "Core Network"
            # dropdown = self.web.browser_xpathclick('//select[@id="ddlWorkType0"]')
            # option_value = self.web.driver.find_element(
            #     By.XPATH,
            #     '//select[@id="ddlWorkType0"]//option[contains(text(), "Core Network")]'
            # ).get_attribute("value")
            # Select(dropdown).select_by_value(option_value)
            # time.sleep(2)
            # in ng-date-picker, clear and input plan start.
            input_box = WebDriverWait(self.web.driver, 10).until(EC.element_to_be_clickable((
                By.XPATH, '//input[@id="txtPlanStartDate0"]')))
            input_box.clear()
            input_box.send_keys(self.plan_start)
            self.web.browser_scrolldown()
            time.sleep(1)
            # click button "Create Work Order", long wait until
            self.web.browser_xpathclick('//button[contains(text(), "Create Work Order")]')
            time.sleep(2)
            # click button confirmation_submit
            self.web.browser_xpathclick('//button[@id="confirmation_submit"]')
            time.sleep(9)
            step_nth[1] = True
            
            # record results into excel
            result = WebDriverWait(self.web.driver, 10).until(EC.presence_of_element_located((
                By.XPATH, '//*[@id="success_s_modal"]//span'))).get_property("innerText")
            ref_no = extract_ref_no(str(result))
            logging.debug(f'Extract Ref.No.: {ref_no} from {result}')
            if ref_no:
                project_excel.loc[order, 'Ref. no.'].value = ref_no
                step_nth[2] = True

            # click button "Create Work Order", long wait until
            self.web.browser_xpathclick('//button[contains(text(), "OK")]')
            time.sleep(6)
            step_nth[3] = True

        except (NoSuchElementException, TimeoutException) as e:
            logging.info(e)
        finally:
            logmsg = f'{project_id},{proms_site},{proms_node},{self.plan_start},' + \
                     ','.join([str(n) for n in step_nth]) + f',{ref_no}'
            logging.info(logmsg)

    def demo_create_wo(self, order, project_excel):
        # 'ProjectID,PromsSite,PromsNode,PlanStart,gotTemplate,submittedWO,foundRefNo,finalOK'
        project_id = project_excel.loc[order, 'Project ID'].value
        proms_site = project_excel.loc[order, 'Proms Site'].value
        proms_node = project_excel.loc[order, 'Proms Node'].value
        step_nth = [False] * 4
        try:
            # click side menu "Work Order", wait for page load
            self.web.browser_xpathclick('//span[contains(text(), "Work Order")]/../..//a')
            time.sleep(2)

            if project_id and proms_site and proms_node:
                step_nth[0] = True
            project_excel.loc[order, 'Ref. no.'].value = project_id
            step_nth[2] = True

        except (NoSuchElementException, TimeoutException) as e:
            logging.info(e)
        finally:
            logmsg = f'{project_id},{proms_site},{proms_node},{self.plan_start},' + \
                     ','.join([str(n) for n in step_nth])
            logging.info(logmsg)

    def run_automation(self):
        """ Command when click Run button """
        current_text = self.filelistBox.get("0.0", "end")
        self.projects = entry_to_list(current_text)
        current_text = self.entryDate.get()
        if current_text:
            self.plan_start = current_text
        # freeze main app
        self.filelistBox.configure(state='disabled')
        self.runButton.configure(state='disabled')
        self.addButton.configure(state='disabled')
        self.web = BROWSER()

        self.web.driver.get(config.get("WEBSITE").get("url"))
        # ============================ Login ===================================
        # input username
        self.web.browser_input('//input[@id="username"]', config.get("WEBSITE").get("user"))

        # input password
        self.web.browser_input('//input[@id="password"]', config.get("WEBSITE").get("pass"))

        # ============================ reCaptcha ===============================
        # tick recaptcha checkbox (switch to recaptcha iframe)
        WebDriverWait(self.web.driver, 10).until(EC.frame_to_be_available_and_switch_to_it((
            By.XPATH, '//iframe[@title="reCAPTCHA"]')))
        WebDriverWait(self.web.driver, 30).until(EC.element_to_be_clickable((
            By.XPATH, '//span[@id="recaptcha-anchor"]'))).click()
        # >> manually choose recaptcha pixel
        # >> after success recaptcha, switch back to default frame
        time.sleep(config.get('Preferences', {}).get('captcha-wait', 20))
        self.web.driver.switch_to.default_content()
        time.sleep(3)

        # click button sign-in
        self.web.browser_xpathclick('//button[@id="btn_signin"]')
        time.sleep(3)

        # ============================ Work Order ===============================
        logging.info('### Automation start..')
        loghead = 'ProjectID,PromsSite,PromsNode,PlanStart,gotTemplate,submittedWO,foundRefNo,finalOK,resultRefNo'
        logging.info(loghead)
        # click side menu "Work Order", wait for page load
        self.web.browser_xpathclick('//span[contains(text(), "Work Order")]/../..//a')
        time.sleep(2)

        # for each project
        for project_file in self.projects:

            from styleframe import StyleFrame
            sf = StyleFrame.read_excel(project_file, sheet_name='WO request', read_style=True)

            # for each WO in a project
            for wo in range(len(sf)):
                wo_func = {'test': self.demo_create_wo, 'real': self.auto_create_wo}
                create_wo = wo_func[config.get('RUNMODE', 'test')]
                create_wo(order=wo, project_excel=sf)

            outfile = os.path.basename(project_file)
            if not os.path.isdir('output'):
                os.mkdir('output')
            xlwriter = sf.to_excel(f'output/{outfile}', sheet_name='WO request')
            xlwriter.close()

        # =========================== Close Automation ===========================
        # Log out
        self.web.browser_xpathclick('//span[contains(text(), "Sign Out")]/../../a')
        time.sleep(2)
        self.web.driver.quit()

        # unfreeze main app
        self.filelistBox.configure(state='normal')
        self.runButton.configure(state='normal')
        self.addButton.configure(state='normal')
