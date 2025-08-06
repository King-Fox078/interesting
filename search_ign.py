import tkinter as tk
from tkinter import scrolledtext
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import threading
import time

def normalize_text(text):
    return text.replace(" ", "").lower()

def search_and_get_scores(keyword, output_widget):
    options = Options()
    options.add_argument('--headless')  # 无头模式，不弹浏览器
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    def log(text):
        output_widget.insert(tk.END, text + "\n")
        output_widget.see(tk.END)

    try:
        search_url = f"https://www.ign.com.cn/se/?q={keyword}"
        log(f"打开搜索页面：{search_url}")
        driver.get(search_url)
        time.sleep(3)

        last_height = driver.execute_script("return document.body.scrollHeight")
        while True:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height

        log("搜索结果加载完毕，开始查找评分...")

        captions = driver.find_elements(By.CLASS_NAME, "caption")
        log(f"共检测到 {len(captions)} 条评分内容。")

        normalized_keyword = normalize_text(keyword)
        found_any = False

        for i, caption in enumerate(captions, 1):
            text = caption.text.strip()
            normalized_text = normalize_text(text)
            if normalized_keyword in normalized_text:
                parent_a = caption.find_element(By.XPATH, "./ancestor::a")
                link = parent_a.get_attribute("href")

                log(f"\n✅ 找到匹配评分（第 {i} 条）：")
                log(f"📝 内容：{text}")
                log(f"🔗 链接：{link}")
                found_any = True

        if not found_any:
            log(f"❌ 没有找到包含关键词“{keyword}”的评分内容。")

    except Exception as e:
        log(f"程序出错：{e}")

    finally:
        driver.quit()

def on_search():
    keyword = entry.get().strip()
    if not keyword:
        text_output.insert(tk.END, "游戏名不能为空！\n")
        return
    # 在新线程中运行，防止界面卡死
    threading.Thread(target=search_and_get_scores, args=(keyword, text_output), daemon=True).start()

# GUI界面
root = tk.Tk()
root.title("IGN游戏评分爬取器")

tk.Label(root, text="请输入游戏名称：").pack(padx=10, pady=5)
entry = tk.Entry(root, width=30)
entry.pack(padx=10)

btn = tk.Button(root, text="搜索评分", command=on_search)
btn.pack(pady=10)

text_output = scrolledtext.ScrolledText(root, width=60, height=20)
text_output.pack(padx=10, pady=10)

root.mainloop()
