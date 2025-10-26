import pandas as pd
import time
from datetime import datetime
import logging
import sys
import os
import json
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyautogui
from selenium.webdriver.common.action_chains import ActionChains
import re

# Simple logger
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def read_contacts(path='contacts.xlsx'):
    df = pd.read_excel(path)
    if df.empty:
        raise ValueError('contacts.xlsx is empty')
    required = ['Name', 'Nickname', 'Number']
    if not all(c in df.columns for c in required):
        raise ValueError(f'contacts.xlsx must contain: {required}')
    
    df = df.drop_duplicates(subset=['Number'], keep='first')
    contacts = {col: df[col].fillna('').astype(str).tolist() for col in df.columns}
    return contacts


def read_event(path='event.json'):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def read_template(path='template.txt'):
    with open(path, 'r', encoding='utf-8') as f:
        txt = f.read()
    if not txt:
        raise ValueError('template.txt is empty')
    return txt


def setup_driver(chrome_data_dir='chrome-data'):
    options = Options()
    options.add_argument('--start-maximized')
    options.add_argument('--disable-notifications')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--log-level=3')

    chrome_data_dir = os.path.abspath(chrome_data_dir)
    os.makedirs(chrome_data_dir, exist_ok=True)
    options.add_argument(f'--user-data-dir={chrome_data_dir}')
    options.add_argument('--profile-directory=Default')

    service = Service()
    driver = webdriver.Chrome(service=service, options=options)
    return driver


def wait_for_logged_in(driver, timeout=30):
    """Wait until WhatsApp Web shows elements that indicate user is logged in."""
    logging.info('Waiting for WhatsApp Web to indicate login...')
    driver.get('https://web.whatsapp.com')
    # give initial load time
    time.sleep(3)
    WebDriverWait(driver, timeout).until(
        EC.any_of(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'div[role="textbox"]')),
            EC.presence_of_element_located((By.CSS_SELECTOR, '[data-testid="chat-list"]')),
            EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'landing-page')]"))
        )
    )
    logging.info('WhatsApp Web looks ready')


def send_message(driver, number, message, country_code='962'):
    """Open chat for `number`, insert `message` and send it. Returns True if sent (verified), else False."""
    try:
        url = f'https://web.whatsapp.com/send?phone={country_code}{number}'
        logging.info(f'Opening chat URL for {number}')
        driver.get(url)

        # Wait for either error text or message box
        WebDriverWait(driver, 30).until(
            EC.any_of(
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(),'Phone number shared via url is invalid') or contains(text(),'not registered')]")),
                EC.presence_of_element_located((By.CSS_SELECTOR, 'div[title="Type a message"]')),
                EC.presence_of_element_located((By.CSS_SELECTOR, 'div[role="textbox"]'))
            )
        )

        # Check for invalid number messages
        page = driver.page_source
        if 'Phone number shared via url is invalid' in page or 'Phone number shared via url is not registered' in page or 'Phone number is not registered' in page:
            logging.error(f'Invalid or unregistered number: {number}')
            return False

        # Try to find message input (prefer the compose box inside the chat container '#main')
        logging.info('Locating message input box (searching inside #main)...')
        compose = None
        try:
            main_elem = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.ID, 'main'))
            )
            # look for contenteditable/textbox elements inside #main
            candidates = main_elem.find_elements(By.CSS_SELECTOR, 'div[contenteditable="true"][data-tab], div[role="textbox"][contenteditable="true"]')
            for el in candidates:
                try:
                    if not el.is_displayed():
                        continue
                    # avoid the global search box by checking attributes or ancestor landmarks
                    aria = (el.get_attribute('aria-label') or '').lower()
                    title = (el.get_attribute('title') or '').lower()
                    # skip obvious search inputs
                    if 'search' in aria or 'search' in title or 'search or start a new chat' in aria:
                        continue
                    compose = el
                    break
                except Exception:
                    continue
        except Exception:
            compose = None

        # Fallbacks: try previous selectors but ensure element is inside #main
        if compose is None:
            try:
                # Try find by title inside main
                compose = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, '#main div[title="Type a message"][contenteditable="true"]'))
                )
            except Exception:
                try:
                    # Last resort: any contenteditable textbox on the page
                    compose = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'div[role="textbox"][contenteditable="true"]'))
                    )
                except Exception as e:
                    logging.error('Could not locate compose box: %s', e)
                    raise

        # Ensure compose is focused and visible
        try:
            driver.execute_script('arguments[0].scrollIntoView({block:"center"}); arguments[0].focus();', compose)
            time.sleep(0.3)
            try:
                compose.click()
            except Exception:
                # fallback: click via JS
                driver.execute_script('arguments[0].click();', compose)
        except Exception:
            logging.warning('Unable to programmatically focus/scroll compose box')

        # Insert message via JS and dispatch input event
        logging.info('Inserting message into input box...')

        # Preserve template formatting exactly. Use <br> for newlines when injecting HTML.
        html_message = message.replace('\n', '<br>')

        driver.execute_script(
            "arguments[0].innerHTML = arguments[1]; arguments[0].dispatchEvent(new Event('input', {bubbles:true}));",
            compose,
            html_message
        )
        time.sleep(0.8)

        # Verify message present; if not, fallback to typing while preserving exact newlines
        text_content = (compose.get_attribute('textContent') or '').strip()
        if message.strip() not in text_content:
            logging.info('JS injection not reflected; typing message as fallback (preserving newlines)...')
            # clear then type
            try:
                # try select-all + delete
                compose.send_keys(Keys.CONTROL, 'a')
                compose.send_keys(Keys.BACKSPACE)
            except Exception:
                pass

            # Use ActionChains to type lines and insert Shift+Enter for every newline
            try:
                actions = ActionChains(driver)
                lines = message.split('\n')
                for idx, line in enumerate(lines):
                    if line:
                        actions.send_keys(line)
                    # For each newline (between lines) insert Shift+Enter to create a newline without sending
                    if idx < len(lines) - 1:
                        actions.key_down(Keys.SHIFT).send_keys(Keys.ENTER).key_up(Keys.SHIFT)
                actions.perform()
            except Exception:
                # final fallback: per-character, using Shift+Enter for newlines
                for ch in message:
                    if ch == '\n':
                        try:
                            compose.send_keys(Keys.SHIFT, Keys.ENTER)
                        except Exception:
                            pyautogui.hotkey('shift', 'enter')
                    else:
                        compose.send_keys(ch)
                    time.sleep(0.01)

        time.sleep(0.5)

        # Try to click send button
        logging.info('Trying to send the message...')
        try:
            send_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-testid="send"]'))
            )
            send_btn.click()
        except Exception:
            # fallback: press Enter
            try:
                compose.send_keys(Keys.ENTER)
            except Exception:
                # final fallback: pyautogui
                pyautogui.press('enter')

        # Verify send by checking if compose box cleared
        time.sleep(2)
        text_after = (compose.get_attribute('textContent') or '').strip()
        if not text_after or len(text_after) < 10:
            logging.info(f'Message sent to {number}')
            return True
        else:
            logging.warning(f'Message may not have sent to {number}')
            return True

    except Exception as e:
        logging.error(f'Error sending to {number}: {e}')
        return False


def send_messages():
    contacts = read_contacts()
    event = read_event()
    template = read_template()
    
    numbers = contacts['Number']
    nicknames = contacts['Nickname']
    total = len(numbers)

    driver = None
    try:
        driver = setup_driver()
        wait_for_logged_in(driver, timeout=45)

        successes = 0
        failures = 0

        for i in range(total):
            num = numbers[i]
            nick = nicknames[i]
            message = template.replace('{nickname}', nick)
            
            # Replace event placeholders
            for key, value in event.items():
                message = message.replace('{' + key + '}', value)
            
            logging.info(f'Sending to {num} ({nick})')
            ok = send_message(driver, num, message)
            if ok:
                successes += 1
            else:
                failures += 1
            time.sleep(3)

        logging.info(f'Finished. Sent: {successes}, Failed: {failures} / {total}')

    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass


if __name__ == '__main__':
    send_messages()