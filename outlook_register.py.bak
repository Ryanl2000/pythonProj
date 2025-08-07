#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import time
import random
import string
import logging
import traceba            # 等待页面完全加载
            self.driver.execute_script("window.scrollTo(0, 0);")  # 滚动到页面顶部
            time.sleep(2)  # 等待页面稳定
            
            # 检查页面是否完全加载
            try:
                self.wait.until(lambda driver: driver.execute_script("return document.readyState") == "complete")
                logging.info("页面加载完成")
            except TimeoutException:
                logging.warning("等待页面加载完成超时")
# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('outlook_register.log'),
        logging.StreamHandler()
    ]
)
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from fake_useragent import UserAgent

class OutlookRegistration:
    def __init__(self):
        self.ua = UserAgent()
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument(f'user-agent={self.ua.random}')
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        chrome_options.add_argument('--disable-notifications')
        chrome_options.add_argument('--start-maximized')  # 最大化窗口
        chrome_options.add_argument('--disable-popup-blocking')  # 禁用弹窗拦截
        chrome_options.add_argument('--disable-infobars')  # 禁用信息栏
        chrome_options.add_argument('--disable-extensions')  # 禁用扩展
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        chrome_options.add_experimental_option('prefs', {
            'profile.default_content_setting_values': {
                'notifications': 2,
                'popups': 2
            }
        })
        # chrome_options.add_argument('--headless')  # 无界面模式，取消注释以启用
        
        try:
            # 直接使用系统安装的Chrome浏览器
            self.driver = webdriver.Chrome(options=chrome_options)
            self.wait = WebDriverWait(self.driver, 60)  # 设置默认等待时间为60秒
            self.driver.set_page_load_timeout(90)  # 设置页面加载超时时间为90秒
            self.driver.set_script_timeout(60)  # 设置脚本超时时间为60秒
            # 设置隐式等待时间
            self.driver.implicitly_wait(30)
            logging.info("成功初始化Chrome浏览器")
        except Exception as e:
            logging.error(f"初始化Chrome浏览器失败: {str(e)}")
            logging.info("尝试使用ChromeDriver Manager...")
            try:
                from selenium.webdriver.chrome.service import Service as ChromeService
                from webdriver_manager.chrome import ChromeDriverManager
                service = ChromeService(ChromeDriverManager(version="114.0.5735.90").install())
                self.driver = webdriver.Chrome(service=service, options=chrome_options)
                self.wait = WebDriverWait(self.driver, 10)
                logging.info("使用ChromeDriver Manager成功初始化浏览器")
            except Exception as e:
                logging.error(f"使用ChromeDriver Manager初始化失败: {str(e)}")
                raise

    def generate_random_info(self):
        # 生成随机用户名
        username = ''.join(random.choices(string.ascii_lowercase, k=8))
        # 生成随机密码
        password = ''.join(random.choices(string.ascii_letters + string.digits + '@#$%', k=12))
        return username, password

    def handle_popups(self):
        """处理可能出现的弹窗和遮罩"""
        try:
            # 检查并关闭cookie提示
            cookie_selectors = [
                "button[aria-label='Accept']",
                "button[data-tid='cookie-banner-close']",
                "#cookie-banner button",
                "button.cookie-accept"
            ]
            for selector in cookie_selectors:
                try:
                    cookie_button = self.wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                    )
                    cookie_button.click()
                    logging.info("关闭cookie提示...")
                    break
                except:
                    continue

            # 检查并关闭其他可能的弹窗
            close_selectors = [
                "button.close-button",
                "button[aria-label='Close']",
                ".modal-close",
                ".dialog-close"
            ]
            for selector in close_selectors:
                try:
                    close_button = self.driver.find_element(By.CSS_SELECTOR, selector)
                    close_button.click()
                    logging.info("关闭弹窗...")
                except:
                    continue

        except Exception as e:
            logging.warning(f"处理弹窗时发生错误: {str(e)}")

    def register(self):
        try:
            # 访问Outlook注册页面
            logging.info("正在访问Outlook注册页面...")
            self.driver.get('https://outlook.live.com/owa/?nlp=1&signup=1')
            
            # 等待页面加载，增加等待时间并添加重试机制
            max_retries = 3
            
            # 确保页面完全加载
            self.driver.execute_script("window.scrollTo(0, 0);")  # 滚动到页面顶部
            time.sleep(2)  # 等待页面稳定
            
            # 处理可能出现的弹窗
            self.handle_popups()
            for retry in range(max_retries):
                try:
                    logging.info(f"等待页面加载...（尝试 {retry + 1}/{max_retries}）")
                    # 等待页面完全加载
                    self.wait.until(lambda driver: driver.execute_script("return document.readyState") == "complete")
                    logging.info("页面加载完成")
                    
                    # 等待页面渲染和动画完成
                    time.sleep(5)  # 增加等待时间，确保页面完全渲染
                    
                    # 尝试等待任何加载指示器消失
                    try:
                        self.wait.until(lambda d: not d.find_elements(By.CSS_SELECTOR, ".loading, .loader, .spinner"))
                        logging.info("加载指示器已消失")
                    except:
                        logging.info("没有发现加载指示器")
                    try:
                        # 尝试多种可能的选择器和方法
                        selectors = [
                            "button[data-text='同意并继续']",
                            "button[data-text='Agree and continue']",  # 英文版
                            "button[type='submit']",
                            "button.ms-Button--primary",
                            "button[class*='primary']",
                            "button.ms-Button",
                            "button[role='button']",
                            "input[type='submit']",
                            "button",  # 尝试查找所有按钮
                            "button:not([disabled])",  # 查找所有未禁用的按钮
                            "[role='button']"  # 查找所有具有button角色的元素
                        ]
                        
                        # 先尝试直接查找和点击
                        for selector in selectors:
                            try:
                                logging.info(f"尝试选择器: {selector}")
                                agree_button = self.wait.until(
                                    EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                                )
                                # 确保元素在视图中且可点击
                                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", agree_button)
                                time.sleep(2)
                                
                                # 尝试多种点击方法
                                try:
                                    agree_button.click()
                                except:
                                    try:
                                        self.driver.execute_script("arguments[0].click();", agree_button)
                                    except:
                                        # 使用ActionChains尝试点击
                                        from selenium.webdriver.common.action_chains import ActionChains
                                        ActionChains(self.driver).move_to_element(agree_button).click().perform()
                                
                                logging.info("点击同意并继续按钮...")
                                
                                # 等待按钮点击后的变化
                                time.sleep(3)
                                try:
                                    # 检查按钮是否消失或变为禁用状态
                                    if not agree_button.is_displayed() or not agree_button.is_enabled():
                                        logging.info("按钮已被点击并发生变化")
                                        break
                                except:
                                    # 如果无法访问元素，可能说明页面已经开始转换
                                    logging.info("无法访问按钮，可能已完成点击")
                                    break
                            except Exception as e:
                                logging.debug(f"选择器 {selector} 失败: {str(e)}")
                                continue
                        
                        # 如果直接点击失败，尝试查找包含特定文本的按钮
                        try:
                            buttons = self.driver.find_elements(By.TAG_NAME, 'button')
                            for button in buttons:
                                try:
                                    button_text = button.text.lower()
                                    if '同意' in button_text or 'agree' in button_text or '继续' in button_text or 'continue' in button_text:
                                        self.driver.execute_script("arguments[0].scrollIntoView(true);", button)
                                        time.sleep(1)
                                        button.click()
                                        logging.info(f"通过文本内容找到并点击按钮: {button_text}")
                                        break
                                except:
                                    continue
                        except Exception as e:
                            logging.debug(f"通过文本内容查找按钮失败: {str(e)}")
                        
                        # 如果上述方法都失败，尝试更多的JavaScript方法
                        js_clicks = [
                            """
                            var buttons = document.getElementsByTagName('button');
                            for(var i = 0; i < buttons.length; i++) {
                                if(buttons[i].innerText.toLowerCase().includes('同意') || 
                                   buttons[i].innerText.toLowerCase().includes('agree') ||
                                   buttons[i].innerText.toLowerCase().includes('继续') ||
                                   buttons[i].innerText.toLowerCase().includes('continue')) {
                                    buttons[i].click();
                                    return true;
                                }
                            }
                            return false;
                            """,
                            "document.querySelector('button[type=\"submit\"]').click();",
                            "document.querySelector('.ms-Button--primary').click();",
                            "document.querySelector('button[role=\"button\"]').click();",
                            """
                            var observer = new MutationObserver(function(mutations) {
                                mutations.forEach(function(mutation) {
                                    if (mutation.addedNodes.length) {
                                        var buttons = document.getElementsByTagName('button');
                                        for(var i = 0; i < buttons.length; i++) {
                                            if(buttons[i].offsetParent !== null && 
                                               (buttons[i].innerText.toLowerCase().includes('同意') || 
                                                buttons[i].innerText.toLowerCase().includes('agree'))) {
                                                buttons[i].click();
                                                observer.disconnect();
                                                return;
                                            }
                                        }
                                    }
                                });
                            });
                            observer.observe(document.body, { childList: true, subtree: true });
                            """
                        ]
                        
                        for js_click in js_clicks:
                            try:
                                self.driver.execute_script(js_click)
                                logging.info("使用JavaScript点击同意并继续按钮...")
                                break
                            except Exception as e:
                                continue
                                
                    except Exception as e:
                        logging.warning(f"所有点击方法都失败: {str(e)}")
                    
                    time.sleep(3)  # 等待页面过渡
                    
                    # 等待邮箱输入框出现
                    try:
                        email_input = self.wait.until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[type="email"]'))
                        )
                    except TimeoutException:
                        # 如果超时，尝试刷新页面
                        logging.info("等待邮箱输入框超时，尝试刷新页面...")
                        self.driver.refresh()
                        time.sleep(5)
                        email_input = self.wait.until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[type="email"]'))
                        )
                    break
                except TimeoutException:
                    if retry < max_retries - 1:
                        logging.warning(f"页面加载超时，正在重试...")
                        self.driver.refresh()
                        time.sleep(2)
                    else:
                        raise

            # 等待页面重新加载完成
            time.sleep(5)
            
            # 生成随机用户信息
            username, password = self.generate_random_info()
            
            # 输入邮箱地址
            logging.info("等待邮箱输入框...")
            email_input = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'input[type="email"]'))
            )
            
            # 确保输入框是可交互的
            self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[type="email"]')))
            
            logging.info("输入邮箱地址...")
            max_retries = 3
            for retry in range(max_retries):
                try:
                    email_input.clear()
                    time.sleep(1)
                    
                    # 先尝试直接输入
                    email_input.send_keys(username)
                    time.sleep(2)
                    
                    # 检查输入值
                    actual_value = email_input.get_attribute('value')
                    if actual_value == username:
                        logging.info(f"成功设置邮箱地址: {actual_value}")
                        break
                        
                    logging.info(f"重试 {retry + 1}: 当前值为 '{actual_value}'，期望值为 '{username}'")
                    
                    # 如果直接输入失败，使用JavaScript方式
                    self.driver.execute_script(f"arguments[0].value = '{username}';", email_input)
                    # 触发所有可能的事件
                    self.driver.execute_script("""
                        var input = arguments[0];
                        input.dispatchEvent(new Event('input', { bubbles: true }));
                        input.dispatchEvent(new Event('change', { bubbles: true }));
                        input.dispatchEvent(new Event('blur', { bubbles: true }));
                        input.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', bubbles: true }));
                    """, email_input)
                    time.sleep(2)
                    
                    # 再次检查
                    actual_value = email_input.get_attribute('value')
                    if actual_value == username:
                        logging.info(f"成功设置邮箱地址: {actual_value}")
                        break
                    
                except Exception as e:
                    logging.warning(f"输入邮箱地址时出错 (尝试 {retry + 1}/{max_retries}): {str(e)}")
                    if retry == max_retries - 1:
                        raise
                
            # 再次验证
            actual_value = email_input.get_attribute('value')
            logging.info(f"设置的邮箱地址: {actual_value}")
            
            # 如果还是空的，尝试最后的方法
            if not actual_value:
                logging.info("尝试模拟人工输入...")
                email_input.clear()
                for char in username:
                    email_input.send_keys(char)
                    time.sleep(0.1)
            
            # 等待并点击"下一个"按钮
            try:
                # 等待页面更新
                time.sleep(2)
                logging.info("查找下一个按钮...")
                
                # 使用更精确的XPath选择器
                try:
                    # 获取所有按钮并打印其文本内容以便调试
                    buttons = self.driver.find_elements(By.TAG_NAME, "button")
                    for btn in buttons:
                        try:
                            logging.info(f"发现按钮: {btn.text}")
                        except:
                            continue
                    
                    # 使用更精确的XPath选择器
                    xpath_selectors = [
                        "//button[normalize-space(.)='下一个']",  # 精确匹配
                        "//button[text()='下一个']",  # 完全匹配
                        "//button[.='下一个']",  # 另一种完全匹配
                        "//button[contains(normalize-space(.), '下一个')]"  # 包含匹配，但会去除多余空格
                    ]
                    
                    for xpath in xpath_selectors:
                        try:
                            next_button = self.driver.find_element(By.XPATH, xpath)
                            if next_button and next_button.is_displayed() and next_button.is_enabled():
                                logging.info(f"通过XPath找到按钮: {next_button.text}")
                                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
                                time.sleep(1)
                                next_button.click()
                                logging.info("点击成功")
                                time.sleep(3)
                                break
                        except:
                            continue
                except Exception as e:
                    logging.debug(f"XPath查找失败: {str(e)}")
                    
                # 如果XPath方法失败，尝试其他方法
                if next_button is None:
                    logging.info("尝试使用JavaScript查找按钮...")
                    # 使用JavaScript查找按钮
                    next_button = self.driver.execute_script("""
                        // 查找所有按钮
                        var buttons = document.getElementsByTagName('button');
                        // 遍历查找文本完全匹配的按钮
                        for (var i = 0; i < buttons.length; i++) {
                            var btnText = buttons[i].textContent.trim();
                            if (btnText === '下一个') {
                                return buttons[i];
                            }
                        }
                        return null;
                    """)
                    
                    if next_button:
                        logging.info("通过JavaScript找到下一个按钮")
                    else:
                        # 如果还是没找到，尝试使用CSS选择器
                        css_selectors = [
                            'button[data-text="下一个"]',
                            'button[aria-label="下一个"]',
                            'button[role="button"]:not([disabled])',
                            'button.ms-Button--primary:not([disabled])'
                        ]
                        
                        for selector in css_selectors:
                            try:
                                next_button = self.wait.until(
                                    EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                                )
                                if next_button and next_button.text.strip() == '下一个':
                                    logging.info(f"通过CSS选择器找到下一个按钮: {selector}")
                                    break
                                next_button = None
                            except:
                                continue
                # 如果通过其他方法都没找到按钮，尝试常用的CSS选择器
                css_selectors = [
                    'button[type="submit"]',
                    'button.ms-Button--primary',
                    'button[data-text="下一个"]',
                    'button.next-button'
                ]
                
                for selector in css_selectors:
                    try:
                        next_button = self.wait.until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                        )
                        if next_button and next_button.text.strip() == '下一个':
                            logging.info(f"找到下一步按钮: {selector}")
                            break
                        next_button = None
                    except:
                        continue
                
                if next_button is None:
                    # 最后尝试通过精确的文本匹配
                    buttons = self.driver.find_elements(By.TAG_NAME, 'button')
                    for button in buttons:
                        try:
                            button_text = button.text.strip()
                            if button_text == '下一个':
                                next_button = button
                                logging.info("通过精确文本匹配找到下一个按钮")
                                break
                        except:
                            continue
                            
                    # 如果还是没找到，记录所有按钮的文本以便调试
                    if next_button is None:
                        logging.info("未找到文本为'下一个'的按钮，当前页面上的按钮文本：")
                        for button in buttons:
                            try:
                                logging.info(f"按钮文本: '{button.text}'")
                            except:
                                continue
                
                if next_button is None:
                    raise Exception("无法找到下一步按钮")
                
                # 确保按钮可见
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
                time.sleep(2)
                
                # 尝试多种点击方法
                try:
                    next_button.click()
                except:
                    try:
                        self.driver.execute_script("arguments[0].click();", next_button)
                    except:
                        from selenium.webdriver.common.action_chains import ActionChains
                        ActionChains(self.driver).move_to_element(next_button).click().perform()
                
                logging.info("点击下一步...")
                time.sleep(3)  # 等待点击后的页面变化
                
            except Exception as e:
                logging.error(f"无法点击下一步按钮: {str(e)}")
                # 保存当前页面源码以供调试
                with open('page.html', 'w', encoding='utf-8') as f:
                    f.write(self.driver.page_source)
                logging.info("已保存页面源码到page.html")
                raise

            # 等待页面加载完成
            logging.info("等待页面加载完成...")
            time.sleep(5)
            self.wait.until(lambda driver: driver.execute_script("return document.readyState") == "complete")
            
            # 等待并检查密码输入框
            logging.info("等待密码输入框...")
            max_retries = 3
            for retry in range(max_retries):
                try:
                    # 检查URL是否已经变化（通常表示页面已经跳转）
                    current_url = self.driver.current_url
                    if "signup" not in current_url.lower():
                        logging.info("检测到页面跳转，重新尝试点击下一步按钮...")
                        raise Exception("页面未正确跳转")
                    
                    # 等待密码输入框
                    password_input = self.wait.until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'input[type="password"]'))
                    )
                    
                    # 确保输入框可见且可交互
                    self.wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[type="password"]'))
                    )
                    logging.info("找到密码输入框")
                    time.sleep(1)
                    break
                except Exception as e:
                    if retry < max_retries - 1:
                        logging.warning(f"未找到密码输入框，正在重试... ({retry + 1}/{max_retries})")
                        # 尝试重新点击下一步按钮
                        try:
                            next_button = self.wait.until(
                                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[type="submit"]'))
                            )
                            next_button.click()
                            time.sleep(3)
                        except:
                            pass
                    else:
                        raise
            
            # 模拟人工输入密码
            logging.info("输入密码...")
            for char in password:
                password_input.send_keys(char)
                time.sleep(0.1)
            
            # MARK: 密码输入完成
            
            # 继续注册流程
            try:
                # 增加更长的等待时间，确保页面完全加载
                time.sleep(5)
                
                # 查看当前URL以确认我们在正确的页面
                current_url = self.driver.current_url
                logging.info(f"当前页面URL: {current_url}")
                
                # 尝试多种方法找到并点击下一步按钮
                submitted = False
                
                # 1. 尝试直接提交表单
                if not submitted:
                    logging.info("尝试提交表单...")
                    try:
                        form = self.driver.find_element(By.TAG_NAME, 'form')
                        if form:
                            self.driver.execute_script("arguments[0].submit();", form)
                            logging.info("表单提交成功")
                            submitted = True
                    except Exception as e:
                        logging.debug(f"表单提交失败: {str(e)}")
                
                # 2. 如果表单提交失败，尝试点击下一步按钮
                if not submitted:
                    next_button_selectors = [
                        'button[type="submit"]',
                        'input[type="submit"]',
                        '#next',
                        '#submitButton',
                        '.button-primary',
                        'button.ms-Button--primary',
                        'button[data-task="next"]'
                    ]
                
                next_button = None
                for selector in next_button_selectors:
                    try:
                        buttons = self.driver.find_elements(By.CSS_SELECTOR, selector)
                        for button in buttons:
                            if button.is_displayed() and button.is_enabled():
                                next_button = button
                                logging.info(f"找到下一步按钮: {selector}")
                                try:
                                    # 使用JavaScript点击按钮
                                    self.driver.execute_script("arguments[0].click();", button)
                                    logging.info("点击下一步按钮成功")
                                    time.sleep(3)
                                    break
                                except:
                                    pass
                    except Exception as e:
                        logging.debug(f"尝试选择器 {selector} 失败: {str(e)}")
                        continue
                        
                # 如果还没有成功，保存页面源码以供分析
                with open('password_page.html', 'w', encoding='utf-8') as f:
                    f.write(self.driver.page_source)
                logging.info("已保存密码页面源码到password_page.html")
                
                # 等待页面完全加载
                self.wait.until(lambda driver: driver.execute_script("return document.readyState") == "complete")
                time.sleep(3)
                
                # 记录页面上所有表单元素
                logging.info("页面上的表单元素：")
                form_elements = self.driver.find_elements(By.CSS_SELECTOR, 'input, button, select, textarea')
                for elem in form_elements:
                    try:
                        elem_info = f"类型: {elem.get_attribute('type')}, "
                        elem_info += f"ID: {elem.get_attribute('id')}, "
                        elem_info += f"名称: {elem.get_attribute('name')}, "
                        elem_info += f"标签: {elem.get_attribute('aria-label')}"
                        logging.info(elem_info)
                    except:
                        pass
                
                # 尝试多个可能的复选框选择器
                checkbox_selectors = [
                    'input[type="checkbox"]',
                    'input[aria-label*="同意"]',
                    'input[aria-label*="agree"]',
                    'label input[type="checkbox"]',
                    '#terms_checkbox',
                    '.terms-checkbox',
                    '[role="checkbox"]'  # 添加role属性选择器
                ]
                
                checkbox_found = False
                for selector in checkbox_selectors:
                    try:
                        # 设置较短的超时时间，避免在每个选择器上都等待太久
                        temp_wait = WebDriverWait(self.driver, 5)
                        checkbox = temp_wait.until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                        )
                        logging.info(f"找到潜在的复选框元素：{selector}")
                        
                        if checkbox.is_displayed():
                            # 首先尝试常规点击
                            try:
                                checkbox.click()
                                checkbox_found = True
                                logging.info(f"已勾选服务条款（使用选择器：{selector}）...")
                                break
                            except Exception as click_error:
                                logging.debug(f"常规点击失败: {str(click_error)}")
                                # 如果常规点击失败，尝试JavaScript点击
                                self.driver.execute_script("arguments[0].click();", checkbox)
                                checkbox_found = True
                                logging.info(f"已通过JavaScript勾选服务条款（使用选择器：{selector}）...")
                                break
                    except Exception as selector_error:
                        logging.debug(f"选择器 {selector} 未找到有效元素: {str(selector_error)}")
                        continue
                
                if not checkbox_found:
                    logging.warning("未找到常规的复选框，尝试查找label元素...")
                    # 尝试通过文本内容查找label并点击
                    labels = self.driver.find_elements(By.TAG_NAME, "label")
                    for label in labels:
                        try:
                            label_text = label.text.lower()
                            if "同意" in label_text or "agree" in label_text:
                                label.click()
                                logging.info("已通过label文本找到并点击服务条款...")
                                checkbox_found = True
                                break
                        except:
                            continue
                
                time.sleep(2)
            except Exception as e:
                logging.warning(f"处理服务条款复选框时出错: {str(e)}")
            
            # 点击下一步
            next_button = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-task="next"]'))
            )
            time.sleep(1)
            next_button.click()
            logging.info("点击下一步...")

            # 等待并输入姓名
            logging.info("等待姓名输入框...")
            first_name = self.wait.until(
                EC.presence_of_element_located((By.ID, 'FirstName'))
            )
            time.sleep(1)
            
            # 模拟人工输入姓名
            logging.info("输入姓名...")
            for char in "John":
                first_name.send_keys(char)
                time.sleep(0.1)
            
            # 输入姓氏
            last_name = self.driver.find_element(By.ID, 'LastName')
            for char in "Doe":
                last_name.send_keys(char)
                time.sleep(0.1)
            
            # 点击下一步
            time.sleep(1)
            next_button = self.wait.until(
                EC.element_to_be_clickable((By.ID, 'iSignupAction'))
            )
            next_button.click()
            logging.info("点击下一步...")

            # 等待并选择出生日期
            logging.info("等待出生日期选择框...")
            birth_month = self.wait.until(
                EC.presence_of_element_located((By.ID, 'BirthMonth'))
            )
            time.sleep(1)
            
            # 选择月份（使用Select类来处理下拉菜单）
            from selenium.webdriver.support.ui import Select
            Select(birth_month).select_by_visible_text("January")
            logging.info("选择出生月份...")
            time.sleep(0.5)
            
            # 选择日期
            birth_day = self.driver.find_element(By.ID, 'BirthDay')
            Select(birth_day).select_by_value("1")
            logging.info("选择出生日期...")
            time.sleep(0.5)
            
            # 选择年份
            birth_year = self.driver.find_element(By.ID, 'BirthYear')
            Select(birth_year).select_by_value("1990")
            logging.info("选择出生年份...")
            time.sleep(0.5)
            
            # 点击下一步
            next_button = self.wait.until(
                EC.element_to_be_clickable((By.ID, 'iSignupAction'))
            )
            next_button.click()
            logging.info("点击下一步...")

            print(f"注册信息：\n用户名: {username}@outlook.com\n密码: {password}")
            
            # 等待人机验证（需要手动操作）
            input("请完成人机验证后按回车继续...")

            return True, f"{username}@outlook.com", password

        except TimeoutException as e:
            logging.error(f"页面加载超时: {str(e)}")
            logging.debug(traceback.format_exc())
            return False, None, None
        except Exception as e:
            logging.error(f"注册过程中发生错误: {str(e)}")
            logging.debug(traceback.format_exc())
            return False, None, None
        finally:
            try:
                if hasattr(self, 'driver'):
                    self.driver.save_screenshot('error.png')  # 保存错误截图
                    # self.driver.quit()  # 注释掉自动关闭浏览器，方便查看结果
            except Exception as e:
                logging.error(f"关闭浏览器时发生错误: {str(e)}")

def check_chrome_installed():
    """检查Chrome是否已安装"""
    import platform
    system = platform.system()
    
    if system == "Darwin":  # macOS
        chrome_path = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
    elif system == "Windows":
        chrome_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
        if not os.path.exists(chrome_path):
            chrome_path = "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe"
    else:  # Linux
        chrome_path = "/usr/bin/google-chrome"
    
    if not os.path.exists(chrome_path):
        logging.error("未检测到Chrome浏览器，请先安装Chrome浏览器")
        return False
    return True

def main():
    if not check_chrome_installed():
        print("请先安装Chrome浏览器后再运行此脚本")
        return
    
    try:
        bot = OutlookRegistration()
        success, email, password = bot.register()
        
        if success:
            logging.info("注册成功！")
            logging.info(f"邮箱: {email}")
            logging.info(f"密码: {password}")
        else:
            logging.error("注册失败，请重试")
    except Exception as e:
        logging.error(f"程序运行出错: {str(e)}")
        logging.debug(traceback.format_exc())

if __name__ == "__main__":
    main()
