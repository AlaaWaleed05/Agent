from pathlib import Path

path = Path("worker.py")
text = path.read_text(encoding="utf-8")

old = '''STUDENT_DELAY_MIN, STUDENT_DELAY_MAX = 5, 10'''
new = '''# Short, varied pauses between students. These are intentionally bounded so\n# a large batch does not accumulate unnecessary idle time.\nSTUDENT_DELAY_MIN, STUDENT_DELAY_MAX = 4, 8'''
if old not in text:
    raise SystemExit("Could not find STUDENT_DELAY settings")
text = text.replace(old, new, 1)

old = '''def human_type(element, text):\n    element.clear()\n    time.sleep(1)\n    for char in str(text):\n        element.send_keys(char)\n        time.sleep(random.uniform(0.15, 0.4))\n\n\ndef slow_wait(seconds, msg=""):\n'''
new = '''def human_delay(min_seconds, max_seconds, msg=""):\n    seconds = random.uniform(min_seconds, max_seconds)\n    if msg:\n        print(f"    ⏳ {msg} ({seconds:.1f}s)...")\n    time.sleep(seconds)\n\n\ndef human_type(element, text):\n    element.clear()\n    human_delay(0.3, 0.7)\n    for char in str(text):\n        element.send_keys(char)\n        time.sleep(random.uniform(0.05, 0.15))\n\n\ndef slow_wait(seconds, msg=""):\n'''
if old not in text:
    raise SystemExit("Could not find human_type block")
text = text.replace(old, new, 1)

repls = [
('''        wait = WebDriverWait(driver, WAIT_TIME)\n        slow_wait(3, "Loading login page")''', '''        wait = WebDriverWait(driver, WAIT_TIME)\n        human_delay(0.8, 1.5, "Loading login page")'''),
('''        email_field.click()\n        slow_wait(3, "Waiting before email")''', '''        email_field.click()\n        human_delay(0.5, 1.2, "Waiting before email")'''),
('''        password_field = driver.find_element(By.CSS_SELECTOR, "input[type='password']")\n        password_field.click()\n        slow_wait(3, "Waiting before password")''', '''        password_field = wait.until(\n            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='password']"))\n        )\n        password_field.click()\n        human_delay(0.5, 1.3, "Waiting before password")'''),
('''        slow_wait(2, "Waiting before click")\n        login_btn = driver.find_element(By.CSS_SELECTOR, "button[type='submit'], input[type='submit']")\n        login_btn.click()''', '''        human_delay(0.7, 1.4, "Waiting before click")\n        login_btn = wait.until(\n            EC.element_to_be_clickable(\n                (By.CSS_SELECTOR, "button[type='submit'], input[type='submit']")\n            )\n        )\n        login_btn.click()'''),
('''            menu_btn.click()\n            slow_wait(0.8, "Waiting for menu to open")''', '''            menu_btn.click()\n            human_delay(0.5, 1.2, "Waiting for menu to open")'''),
('''            my_apps.click()\n            slow_wait(1, "Waiting for requests page")''', '''            my_apps.click()\n            human_delay(0.7, 1.5, "Waiting for requests page")'''),
('''            driver.get(INBOX_URL)\n            slow_wait(1)''', '''            driver.get(INBOX_URL)\n            human_delay(0.7, 1.5, "Waiting for requests page")'''),
('''        wait = WebDriverWait(driver, WAIT_TIME)\n        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table")))\n        slow_wait(2)''', '''        wait = WebDriverWait(driver, WAIT_TIME)\n        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table")))\n        human_delay(0.5, 1.2, "Reading application status")'''),
('''                slow_wait(random.uniform(STUDENT_DELAY_MIN, STUDENT_DELAY_MAX), "Pause before next student")''', '''                human_delay(STUDENT_DELAY_MIN, STUDENT_DELAY_MAX, "Pause before next student")'''),
('''                slow_wait(random.uniform(STUDENT_DELAY_MIN, STUDENT_DELAY_MAX), "Pause before retry")''', '''                human_delay(STUDENT_DELAY_MIN, STUDENT_DELAY_MAX, "Pause before retry")'''),
]
for old, new in repls:
    if old not in text:
        raise SystemExit("Could not find expected timing block:\n" + old)
    text = text.replace(old, new, 1)

old = '''def selenium_logout(driver):\n    try:\n        wait = WebDriverWait(driver, WAIT_TIME)\n        slow_wait(2)\n        user_menu = wait.until(EC.element_to_be_clickable(\n            (By.CSS_SELECTOR, "[class*='user'], [class*='profile'], [class*='avatar'], [class*='account']")\n        ))\n        user_menu.click()\n        slow_wait(2)\n        logout_btn = wait.until(EC.element_to_be_clickable(\n            (By.XPATH, "//*[contains(text(), 'تسجيل خروج') or contains(text(), 'خروج')]")\n        ))\n        logout_btn.click()\n        slow_wait(2)\n    except Exception:\n        pass\n    finally:\n        clear_session(driver)\n'''
new = '''def selenium_logout(driver):\n    success = False\n    try:\n        wait = WebDriverWait(driver, WAIT_TIME)\n        human_delay(0.8, 1.8, "Preparing logout")\n        user_menu = wait.until(EC.element_to_be_clickable(\n            (By.CSS_SELECTOR, "[class*='user'], [class*='profile'], [class*='avatar'], [class*='account']")\n        ))\n        user_menu.click()\n        human_delay(0.5, 1.2, "Opening account menu")\n        logout_btn = wait.until(EC.element_to_be_clickable(\n            (By.XPATH, "//*[contains(text(), 'تسجيل خروج') or contains(text(), 'خروج')]")\n        ))\n        logout_btn.click()\n        human_delay(0.7, 1.5, "Finishing logout")\n        success = True\n    except Exception as exc:\n        print(f"    ⚠️ Logout failed: {exc}")\n    finally:\n        clear_session(driver)\n    return success\n'''
if old not in text:
    raise SystemExit("Could not find selenium_logout block")
text = text.replace(old, new, 1)

old = '''            finally:\n                \n                if login_confirmed_failed:\n                    \n                    clear_session(driver)\n                else:\n                    \n                    selenium_logout(driver)\n\n    # أي technical error من Selenium ممكن يكون خلّى الـsession غير صالحة،\n    # لذلك نبدأ بمتصفح جديد قبل الطالب التالي.\n                if browser_crashed or is_tech_error:\n                    driver = restart_browser(driver)\n'''
new = '''            finally:\n                if login_confirmed_failed:\n                    clear_session(driver)\n                else:\n                    logout_ok = selenium_logout(driver)\n                    if not logout_ok:\n                        browser_crashed = True\n\n                # A Selenium technical error can leave the current browser\n                # session unusable even when the exception was caught inside\n                # selenium_login()/selenium_get_status(). Start clean.\n                if browser_crashed or is_tech_error:\n                    driver = restart_browser(driver)\n'''
if old not in text:
    raise SystemExit("Could not find first-pass finally block")
text = text.replace(old, new, 1)

old = '''                finally:\n                    \n                    if retry_login_failed:\n                        \n                        clear_session(driver)\n                    else:\n                        selenium_logout(driver)\n\n                    if retry_browser_crashed or retry_is_tech_error:\n                        driver = restart_browser(driver)\n'''
new = '''                finally:\n                    if retry_login_failed:\n                        clear_session(driver)\n                    else:\n                        logout_ok = selenium_logout(driver)\n                        if not logout_ok:\n                            retry_browser_crashed = True\n\n                    if retry_browser_crashed or retry_is_tech_error:\n                        driver = restart_browser(driver)\n'''
if old not in text:
    raise SystemExit("Could not find retry finally block")
text = text.replace(old, new, 1)

path.write_text(text, encoding="utf-8")
print("worker.py timing + browser recovery patch applied")
