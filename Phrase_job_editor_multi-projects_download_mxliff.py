import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import concurrent.futures
import logging
import time
import json
import os
import sys
from datetime import datetime, timezone

# 設置日誌記錄(考慮打包後的路徑)
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
logging.basicConfig(
    filename=os.path.join(BASE_DIR, 'job_status_update.log'),
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# 全域變數
BASE_URL = "https://cloud.memsource.com/web"
API_TOKEN = None
HEADERS = None
CREDENTIALS_FILE = os.path.join(BASE_DIR, "credentials.json")
selected_projects_global = []
selected_target_langs = []  # 新增:儲存選定的目標語言


def save_credentials(username, token, expires):
    """ 儲存帳號、token 和過期時間到檔案(不儲存密碼) """
    try:
        credentials = {
            "username": username,
            "token": token,
            "expires": expires
        }
        with open(CREDENTIALS_FILE, 'w') as f:
            json.dump(credentials, f)
        logging.info("Credentials saved to file")
    except Exception as e:
        logging.error(f"Failed to save credentials: {type(e).__name__} - {str(e)}")
        messagebox.showerror("Error", f"無法儲存憑證: {str(e)}")


def load_credentials():
    """ 從檔案載入帳號、token 和過期時間(不載入密碼) """
    try:
        if os.path.exists(CREDENTIALS_FILE):
            with open(CREDENTIALS_FILE, 'r') as f:
                credentials = json.load(f)
            return credentials
    except (json.JSONDecodeError, IOError) as e:
        logging.error(f"Failed to load credentials: {type(e).__name__} - {str(e)}")
    return None


def is_token_valid(expires):
    """ 檢查 token 是否過期 """
    if not expires:
        return False
    try:
        expire_time = datetime.fromisoformat(expires.replace('Z', '+00:00'))
        current_time = datetime.now(timezone.utc)
        return current_time < expire_time
    except ValueError:
        return False


def login():
    """ 使用帳號和密碼登入 Phrase TMS API,取得 token """
    username = entry_username.get()
    password = entry_password.get()
    if not username or not password:
        messagebox.showerror("Error", "請輸入帳號和密碼")
        return False

    url = f"{BASE_URL}/api2/v1/auth/login"
    payload = {
        "userName": username,
        "password": password
    }
    try:
        response = requests.post(url, json=payload)
        response.raise_for_status()
        data = response.json()
        global API_TOKEN, HEADERS
        API_TOKEN = data["token"]
        HEADERS = {
            "Authorization": f"ApiToken {API_TOKEN}",
            "Content-Type": "application/json"
        }
        save_credentials(username, API_TOKEN, data["expires"])
        entry_project.config(state="normal")
        combo_workflow.config(state="normal")
        combo_status.config(state="normal")
        button_search.config(state="normal")
        button_show_jobs.config(state="normal")
        button_update_status.config(state="normal")
        button_download_bilingual.config(state="normal")  # 新增
        label_login_status.config(text="登入成功", foreground="green")
        logging.info(f"Login successful for user: {username}")
        return True
    except requests.exceptions.HTTPError as http_err:
        messagebox.showerror("Error", f"登入失敗: 帳號密碼錯誤或帳號未啟用")
        logging.error(f"Login failed: {http_err}")
        return False
    except requests.exceptions.RequestException as req_err:
        messagebox.showerror("Error", f"網路錯誤: {req_err}")
        logging.error(f"Login failed: {req_err}")
        return False


def list_projects(project_name=None, client_name=None):
    """查詢包含指定條件的所有專案(自動抓取所有分頁),並取得 targetLangs"""
    global API_TOKEN, HEADERS
    url = f"{BASE_URL}/api2/v1/projects"
    params = {"pageNumber": 0, "pageSize": 50}
    projects_all = []

    if project_name:
        params["name"] = project_name
    if client_name:
        params["clientName"] = client_name

    while True:
        try:
            response = requests.get(url, headers=HEADERS, params=params)
            response.raise_for_status()
            data = response.json()

            # 累積每一頁的結果,並保留 targetLangs 資訊
            for project in data.get("content", []):
                # 確保 targetLangs 資訊被保留
                if 'targetLangs' not in project:
                    project['targetLangs'] = []
                projects_all.append(project)

            total_pages = data.get("totalPages", 1)
            if params["pageNumber"] >= total_pages - 1:
                break
            params["pageNumber"] += 1

        except requests.exceptions.HTTPError as http_err:
            if http_err.response.status_code == 401:
                messagebox.showerror("Error", "Token 已過期,請重新登入")
                API_TOKEN = None
                HEADERS = None
                entry_project.config(state="disabled")
                entry_client.config(state="disabled")
                combo_workflow.config(state="disabled")
                combo_status.config(state="disabled")
                button_search.config(state="disabled")
                button_show_jobs.config(state="disabled")
                button_update_status.config(state="disabled")
                button_download_bilingual.config(state="disabled")
                label_login_status.config(text="未登入", foreground="red")
                raise Exception("Token expired")
            raise Exception(f"Failed to list projects: {http_err}")
        except requests.exceptions.RequestException as e:
            raise Exception(f"Failed to list projects: {e}")

    return projects_all


def list_jobs(project_uid, workflowLevel=None, targetLang=None):
    """ 抓取指定 project 的 jobs,可依 workflowLevel 和 targetLang 篩選
        如果 workflowLevel=None,則抓取所有 jobs
        如果 targetLang 有指定,則只回傳該語系的 jobs
    """
    global API_TOKEN, HEADERS
    jobs_all = []
    page = 0
    while True:
        url = f"{BASE_URL}/api2/v1/projects/{project_uid}/jobs"
        params = {"pageNumber": page}
        if workflowLevel:
            params["workflowLevel"] = workflowLevel
        if targetLang:
            params["targetLang"] = targetLang
        try:
            response = requests.get(url, headers=HEADERS, params=params)
            response.raise_for_status()
            data = response.json()
            jobs = data.get("content", [])
            jobs_all.extend(jobs)
            total_pages = data.get("totalPages", 1)
            if page >= total_pages - 1:
                break
            page += 1
        except requests.exceptions.HTTPError as http_err:
            if http_err.response.status_code == 401:
                messagebox.showerror("Error", "Token 已過期,請重新登入")
                API_TOKEN = None
                HEADERS = None
                entry_project.config(state="disabled")
                combo_workflow.config(state="disabled")
                combo_status.config(state="disabled")
                button_search.config(state="disabled")
                button_show_jobs.config(state="disabled")
                button_update_status.config(state="disabled")
                button_download_bilingual.config(state="disabled")
                label_login_status.config(text="未登入", foreground="red")
                raise Exception("Token expired")
            raise Exception(f"Failed to list jobs: {http_err}")
        except requests.exceptions.RequestException as e:
            raise Exception(f"Failed to list jobs: {e}")
    return jobs_all


def download_bilingual_file(project_uid, job_uids, save_path):
    """ 下載雙語檔案 """
    global API_TOKEN, HEADERS
    url = f"{BASE_URL}/api2/v1/projects/{project_uid}/jobs/bilingualFile"

    # 建立 payload
    jobs_list = [{"uid": uid} for uid in job_uids]
    payload = {"jobs": jobs_list}

    try:
        response = requests.post(url, json=payload, headers=HEADERS)
        response.raise_for_status()

        # 儲存檔案
        with open(save_path, 'wb') as f:
            f.write(response.content)

        return f"成功下載到: {save_path}"
    except requests.exceptions.HTTPError as http_err:
        return f"HTTP error: {http_err}"
    except requests.exceptions.RequestException as req_err:
        return f"Request error: {req_err}"
    except Exception as e:
        return f"Error: {str(e)}"


def select_target_languages():
    """ 顯示語言選擇視窗供使用者複選 """
    if not selected_projects_global:
        messagebox.showerror("Error", "請先選擇專案")
        return

    # 收集所有專案的 targetLangs
    all_langs = set()
    for project in selected_projects_global:
        target_langs = project.get('targetLangs', [])
        all_langs.update(target_langs)

    if not all_langs:
        messagebox.showerror("Error", "所選專案沒有目標語言")
        return

    all_langs = sorted(list(all_langs))

    def on_confirm():
        selected_indices = listbox_langs.curselection()
        if not selected_indices:
            messagebox.showerror("Error", "請選擇至少一個語言")
            return
        global selected_target_langs
        selected_target_langs = [all_langs[i] for i in selected_indices]
        lang_window.destroy()
        label_selected_langs.config(text=f"已選擇語言: {', '.join(selected_target_langs)}")
        logging.info(f"Selected target languages: {selected_target_langs}")

    def on_cancel():
        lang_window.destroy()
        logging.info("Language selection cancelled")

    lang_window = tk.Toplevel(root)
    lang_window.title("選擇目標語言")
    lang_window.geometry("400x400")
    lang_window.transient(root)
    lang_window.grab_set()

    tk.Label(lang_window, text="請選擇要下載的語言(可複選):").pack(pady=10)

    scrollbar = tk.Scrollbar(lang_window)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    listbox_langs = tk.Listbox(lang_window, selectmode=tk.EXTENDED, yscrollcommand=scrollbar.set, height=15, width=50)
    for lang in all_langs:
        listbox_langs.insert(tk.END, lang)
    listbox_langs.pack(pady=10)
    scrollbar.config(command=listbox_langs.yview)

    button_frame = tk.Frame(lang_window)
    button_frame.pack(pady=10)
    tk.Button(button_frame, text="確定", command=on_confirm).pack(side=tk.LEFT, padx=5)
    tk.Button(button_frame, text="取消", command=on_cancel).pack(side=tk.LEFT, padx=5)


def get_unique_filename(directory, filename):
    """ 如果檔名重複，自動加上編號 (1), (2), (3)... """
    base_name, extension = os.path.splitext(filename)
    counter = 1
    new_filename = filename

    while os.path.exists(os.path.join(directory, new_filename)):
        new_filename = f"{base_name} ({counter}){extension}"
        counter += 1

    return new_filename


def download_bilingual_files_by_language():
    """ 根據選定的語言下載雙語檔案 """
    if not API_TOKEN:
        messagebox.showerror("Error", "請先登入")
        return
    if not selected_projects_global:
        messagebox.showerror("Error", "請先選擇專案")
        return
    if not selected_target_langs:
        messagebox.showerror("Error", "請先選擇目標語言")
        return

    # 取得下載模式
    download_mode = download_mode_var.get()

    # 選擇儲存目錄
    save_dir = filedialog.askdirectory(title="選擇儲存目錄")
    if not save_dir:
        return

    workflow_name = combo_workflow.get().strip()

    text_jobs.delete("1.0", tk.END)
    text_jobs.insert(tk.END, f"開始下載雙語檔案 (模式: {download_mode})...\n\n")

    for project in selected_projects_global:
        project_name = project['name']
        project_uid = project['uid']

        text_jobs.insert(tk.END, f"處理專案: {project_name}\n")

        # 確定 workflow level 和取得 workflow abbreviation
        workflow_abbr = ""
        if workflow_name == "No Workflow":
            workflow_level = None
        else:
            if not workflow_name:
                text_jobs.insert(tk.END, f"Skip: 請輸入工作流程名稱\n")
                continue
            workflow_steps = project.get("workflowSteps", [])
            workflow_info = next((w for w in workflow_steps if w["name"] == workflow_name), None)
            if workflow_info is None:
                text_jobs.insert(tk.END, f"Skip: Workflow {workflow_name} not found in project {project_name}\n")
                continue
            workflow_level = workflow_info["workflowLevel"]
            workflow_abbr = workflow_info.get("abbreviation", "")  # 取得縮寫

        try:
            # 為每個選定的語言下載檔案
            for target_lang in selected_target_langs:
                text_jobs.insert(tk.END, f"  處理語言: {target_lang}\n")

                # 取得該語言的 jobs
                jobs = list_jobs(project_uid, workflowLevel=workflow_level, targetLang=target_lang)

                if not jobs:
                    text_jobs.insert(tk.END, f"    沒有找到 {target_lang} 的 jobs\n")
                    continue

                text_jobs.insert(tk.END, f"    找到 {len(jobs)} 個 jobs\n")

                safe_lang = target_lang.replace('/', '_')

                if download_mode == "合併下載":
                    # 合併下載：一個語言一個檔案
                    job_uids = [job['uid'] for job in jobs]

                    # 建立語系資料夾 (加上 workflow 縮寫)
                    if workflow_abbr:
                        folder_name = f"{safe_lang}_{workflow_abbr}"
                    else:
                        folder_name = safe_lang
                    lang_folder = os.path.join(save_dir, folder_name)
                    os.makedirs(lang_folder, exist_ok=True)

                    # 建立檔案名稱: 專案名稱_語系_workflow縮寫.mxliff
                    safe_project_name = "".join(c for c in project_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
                    if workflow_abbr:
                        filename = f"{safe_project_name}_{safe_lang}_{workflow_abbr}.mxliff"
                    else:
                        filename = f"{safe_project_name}_{safe_lang}.mxliff"

                    # 檢查檔名是否重複，如果重複則加上編號
                    filename = get_unique_filename(lang_folder, filename)
                    save_path = os.path.join(lang_folder, filename)

                    # 下載合併的雙語檔案
                    result = download_bilingual_file(project_uid, job_uids, save_path)
                    text_jobs.insert(tk.END, f"    {result}\n")
                    logging.info(f"Project {project_name} - Language {target_lang} (merged): {result}")

                else:  # 單獨下載
                    # 單獨下載：每個 job 一個檔案，使用平行下載加速

                    def download_single_job(job):
                        """下載單一 job 的函數（用於平行處理）"""
                        job_uid = job['uid']
                        job_filename = job.get('filename', 'unknown')

                        # 建立語系資料夾 (加上 workflow 縮寫)
                        if workflow_abbr:
                            folder_name = f"{safe_lang}_{workflow_abbr}"
                        else:
                            folder_name = safe_lang
                        lang_folder = os.path.join(save_dir, folder_name)
                        os.makedirs(lang_folder, exist_ok=True)

                        # 建立檔案名稱: 檔名_語系_workflow縮寫.mxliff
                        # 移除原始檔案的副檔名
                        base_filename = os.path.splitext(job_filename)[0]
                        safe_filename = "".join(
                            c for c in base_filename if c.isalnum() or c in (' ', '-', '_')).rstrip()
                        if not safe_filename:  # 如果過濾後檔名是空的
                            safe_filename = "unnamed"

                        if workflow_abbr:
                            filename = f"{safe_filename}_{safe_lang}_{workflow_abbr}.mxliff"
                        else:
                            filename = f"{safe_filename}_{safe_lang}.mxliff"

                        # 檢查檔名是否重複，如果重複則加上編號
                        filename = get_unique_filename(lang_folder, filename)
                        save_path = os.path.join(lang_folder, filename)

                        # 下載單一 job 的雙語檔案
                        result = download_bilingual_file(project_uid, [job_uid], save_path)
                        return (job_filename, filename, result)

                    # 使用多執行緒平行下載
                    max_workers = min(10, len(jobs))  # 最多同時 10 個下載
                    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
                        future_to_job = {executor.submit(download_single_job, job): job for job in jobs}
                        for future in concurrent.futures.as_completed(future_to_job):
                            try:
                                job_filename, final_filename, result = future.result()
                                text_jobs.insert(tk.END, f"    [{job_filename}] → {final_filename}: {result}\n")
                                text_jobs.see(tk.END)  # 自動捲動到最新訊息
                                root.update()  # 更新 GUI 顯示
                                logging.info(
                                    f"Project {project_name} - Job {job_filename} - Language {target_lang}: {result}")
                            except Exception as e:
                                text_jobs.insert(tk.END, f"    下載失敗: {str(e)}\n")
                                logging.error(f"Download failed: {str(e)}")

            text_jobs.insert(tk.END, f"專案 {project_name} 完成\n\n")

        except Exception as e:
            text_jobs.insert(tk.END, f"Error in project {project_name}: {str(e)}\n\n")
            logging.error(f"Download failed for project {project_name}: {str(e)}")

    text_jobs.insert(tk.END, "所有下載完成!\n")
    messagebox.showinfo("完成", "所有雙語檔案下載完成!")


def update_job_status(project_uid, job_uid, status, retries=3):
    """ 修改單一 job 的 status,包含重試機制 """
    global API_TOKEN, HEADERS
    url = f"{BASE_URL}/api2/v1/projects/{project_uid}/jobs/{job_uid}/setStatus"
    payload = {
        "requestedStatus": status,
        "notifyOwner": True,
        "propagateStatus": True
    }
    session = requests.Session()
    retries = Retry(total=retries, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504])
    session.mount('https://', HTTPAdapter(max_retries=retries))
    try:
        response = session.post(url, headers=HEADERS, json=payload)
        response.raise_for_status()
        return "Status updated successfully"
    except requests.exceptions.HTTPError as http_err:
        if http_err.response.status_code == 401:
            messagebox.showerror("Error", "Token 已過期,請重新登入")
            API_TOKEN = None
            HEADERS = None
            entry_project.config(state="disabled")
            combo_workflow.config(state="disabled")
            combo_status.config(state="disabled")
            button_search.config(state="disabled")
            button_show_jobs.config(state="disabled")
            button_update_status.config(state="disabled")
            button_download_bilingual.config(state="disabled")
            label_login_status.config(text="未登入", foreground="red")
            return "Token expired"
        return f"HTTP error: {http_err} - Response: {response.text if 'response' in locals() else 'No response'}"
    except requests.exceptions.RequestException as req_err:
        return f"Request error: {req_err}"


def select_projects(projects):
    """ 顯示多個專案供使用者選擇(支援多選) """

    def on_confirm():
        selected_indices = listbox.curselection()
        if not selected_indices:
            messagebox.showerror("Error", "請選擇至少一個專案")
            return
        selected_projects = [projects[i] for i in selected_indices]
        global selected_projects_global, selected_target_langs
        selected_projects_global = selected_projects
        if len(selected_projects) == 1:
            project = selected_projects[0]
            label_project_uid.config(text=f"Project UID: {project['uid']}")

            # 自動選擇語系（如果只有單一語系）
            target_langs = project.get('targetLangs', [])
            if len(target_langs) == 1:
                selected_target_langs = target_langs
                label_selected_langs.config(text=f"已自動選擇語言: {target_langs[0]}")
                logging.info(f"Auto-selected single target language: {target_langs[0]}")
            else:
                selected_target_langs = []
                label_selected_langs.config(text="尚未選擇語言")
        else:
            label_project_uid.config(text="Multiple projects selected")
            selected_target_langs = []
            label_selected_langs.config(text="尚未選擇語言")
        project_window.destroy()
        text_jobs.delete("1.0", tk.END)
        text_jobs.insert(tk.END, "Selected projects:\n" + "\n".join([p['name'] for p in selected_projects]))
        logging.info(f"Selected projects: {[p['name'] for p in selected_projects]}")

    def on_cancel():
        project_window.destroy()
        logging.info("Project selection cancelled")

    project_window = tk.Toplevel(root)
    project_window.title("選擇專案")
    project_window.geometry("600x400")
    project_window.transient(root)
    project_window.grab_set()

    tk.Label(project_window, text="找到多個專案,請選擇(可多選):").pack(pady=10)

    scrollbar = tk.Scrollbar(project_window)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    listbox = tk.Listbox(project_window, selectmode=tk.EXTENDED, yscrollcommand=scrollbar.set, height=15, width=80)
    project_options = [
        f"{p['name']} (internalId: {p['internalId']}, Created: {p['dateCreated'][:10]}, Owner: {p['owner']['firstName']} {p['owner']['lastName']})"
        for p in projects
    ]
    for opt in project_options:
        listbox.insert(tk.END, opt)
    listbox.pack(pady=10)
    scrollbar.config(command=listbox.yview)

    button_frame = tk.Frame(project_window)
    button_frame.pack(pady=10)
    tk.Button(button_frame, text="確定", command=on_confirm).pack(side=tk.LEFT, padx=5)
    tk.Button(button_frame, text="取消", command=on_cancel).pack(side=tk.LEFT, padx=5)


def search_project():
    """ 查詢專案,可同時使用專案名稱(多行)與客戶名稱條件 """
    if not API_TOKEN:
        messagebox.showerror("Error", "請先登入")
        return

    # 取得使用者輸入
    project_names_text = entry_project.get("1.0", tk.END).strip()
    client_name = entry_client.get().strip()

    # 將多行專案名稱轉為清單
    project_names = [name.strip() for name in project_names_text.split('\n') if name.strip()]

    if not project_names and not client_name:
        messagebox.showerror("Error", "請輸入至少一個專案名稱或客戶名稱")
        return

    try:
        projects_all = []
        seen_uids = set()

        # 若有專案名稱,逐一查詢
        if project_names:
            for project_name in project_names:
                projects = list_projects(project_name=project_name, client_name=client_name or None)
                for p in projects:
                    if p['uid'] not in seen_uids:
                        seen_uids.add(p['uid'])
                        projects_all.append(p)
        else:
            # 若沒有專案名稱,但有客戶名稱
            projects = list_projects(client_name=client_name)
            for p in projects:
                if p['uid'] not in seen_uids:
                    seen_uids.add(p['uid'])
                    projects_all.append(p)

        # 無結果
        if not projects_all:
            messagebox.showerror("Error", "未找到符合條件的專案")
            return

        # 依建立日期降冪排序
        projects_all.sort(key=lambda x: x['dateCreated'], reverse=True)

        # 多專案 → 顯示選擇視窗
        if len(projects_all) > 1:
            select_projects(projects_all)
        else:
            global selected_projects_global, selected_target_langs
            selected_projects_global = projects_all
            project = projects_all[0]
            label_project_uid.config(text=f"Project UID: {project['uid']}")
            text_jobs.delete("1.0", tk.END)
            text_jobs.insert(tk.END, f"Selected project: {project['name']} (internalId: {project['internalId']})")
            logging.info(
                f"Selected project: {project['name']} (internalId: {project['internalId']}, uid: {project['uid']})")

            # 自動選擇語系（如果只有單一語系）
            target_langs = project.get('targetLangs', [])
            if len(target_langs) == 1:
                selected_target_langs = target_langs
                label_selected_langs.config(text=f"已自動選擇語言: {target_langs[0]}")
                logging.info(f"Auto-selected single target language: {target_langs[0]}")
            else:
                selected_target_langs = []
                label_selected_langs.config(text="尚未選擇語言")

    except Exception as e:
        messagebox.showerror("Error", str(e))
        logging.error(f"Search project failed: {str(e)}")


def show_jobs():
    """ 顯示指定工作流程的任務列表 """
    if not API_TOKEN:
        messagebox.showerror("Error", "請先登入")
        return
    if len(selected_projects_global) > 1:
        messagebox.showerror("Error", "請選擇單一專案來顯示任務")
        return

    workflow_name = combo_workflow.get().strip()
    project = selected_projects_global[0]

    # 檢查是否選擇「No Workflow」
    if workflow_name == "No Workflow":
        workflow_level = None
    else:
        workflow_steps = project.get("workflowSteps", [])
        workflow_level = next((w["workflowLevel"] for w in workflow_steps if w["name"] == workflow_name), None)
        if workflow_level is None and workflow_name:
            messagebox.showerror("Error", f"Workflow {workflow_name} not found")
            return

    try:
        jobs = list_jobs(project['uid'], workflowLevel=workflow_level)
        if not jobs:
            text_jobs.delete("1.0", tk.END)
            text_jobs.insert(tk.END, f"No jobs found")
            return
        jobs_text = "\n".join([f"{j['uid']}: {j['filename']}" for j in jobs])
        text_jobs.delete("1.0", tk.END)
        text_jobs.insert(tk.END, jobs_text)
    except Exception as e:
        messagebox.showerror("Error", str(e))
        logging.error(f"Show jobs failed: {str(e)}")


def update_all_jobs_status():
    """ 批量更新指定工作流程層級的任務狀態 """
    if not API_TOKEN:
        messagebox.showerror("Error", "請先登入")
        return
    if not selected_projects_global:
        messagebox.showerror("Error", "請先選擇專案")
        return

    workflow_name = combo_workflow.get().strip()
    new_status = combo_status.get()

    if not new_status:
        messagebox.showerror("Error", "請選擇狀態")
        return

    text_jobs.delete("1.0", tk.END)

    for project in selected_projects_global:
        text_jobs.insert(tk.END, f"Processing project: {project['name']}\n")

        # 檢查是否選擇「No Workflow」
        if workflow_name == "No Workflow":
            workflow_level = None
        else:
            if not workflow_name:
                text_jobs.insert(tk.END, f"Skip: 請輸入工作流程名稱\n")
                continue
            workflow_steps = project.get("workflowSteps", [])
            workflow_level = next((w["workflowLevel"] for w in workflow_steps if w["name"] == workflow_name), None)
            if workflow_level is None:
                text_jobs.insert(tk.END, f"Skip: Workflow {workflow_name} not found in project {project['name']}\n")
                continue

        try:
            start_time = time.time()
            jobs = list_jobs(project['uid'], workflowLevel=workflow_level)

            if not jobs:
                text_jobs.insert(tk.END, f"No jobs found in project {project['name']}\n")
                continue

            success_count = 0
            fail_count = 0
            batch_size = 50

            def update_single_job(job):
                job_uid = job["uid"]
                job_filename = job.get("filename", "N/A")
                job_workflow_level = job.get("workflowLevel")

                # 如果指定了 workflow level,檢查是否匹配
                if workflow_level is not None and job_workflow_level != workflow_level:
                    return (job_uid, job_filename, f"跳過: 屬於工作流程層級 {job_workflow_level}")

                try:
                    result = update_job_status(project['uid'], job_uid, new_status)
                    return (job_uid, job_filename,
                            result if result == "Status updated successfully" else f"更新失敗: {result}")
                except Exception as e:
                    return (job_uid, job_filename, f"更新失敗: {str(e)}")

            for i in range(0, len(jobs), batch_size):
                batch = jobs[i:i + batch_size]
                max_workers = min(50, len(batch))
                with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
                    future_to_job = {executor.submit(update_single_job, job): job for job in batch}
                    for future in concurrent.futures.as_completed(future_to_job):
                        job_uid, job_filename, result = future.result()
                        text_jobs.insert(tk.END, f"{project['name']} - {job_filename} ({job_uid}) → {result}\n")
                        if result == "Status updated successfully":
                            success_count += 1
                        elif not result.startswith("跳過"):
                            fail_count += 1
                        logging.info(f"Project {project['name']} Job {job_uid} ({job_filename}): {result}")

            elapsed_time = time.time() - start_time
            text_jobs.insert(tk.END,
                             f"Project {project['name']}: 成功更新 {success_count} 個 jobs,失敗 {fail_count} 個,耗時 {elapsed_time:.2f} 秒\n\n")
            logging.info(
                f"Project {project['name']}: Batch update completed: {success_count} successful, {fail_count} failed, took {elapsed_time:.2f} seconds")

        except Exception as e:
            text_jobs.insert(tk.END, f"Error in project {project['name']}: {str(e)}\n")
            logging.error(f"Batch update failed for project {project['name']}: {str(e)}")


def clear_credentials():
    """ 清除儲存的憑證 """
    try:
        if os.path.exists(CREDENTIALS_FILE):
            os.remove(CREDENTIALS_FILE)
        key_file = os.path.join(BASE_DIR, "key.key")
        if os.path.exists(key_file):
            os.remove(key_file)
        entry_username.delete(0, tk.END)
        entry_password.delete(0, tk.END)
        label_login_status.config(text="憑證已清除", foreground="blue")
        logging.info("Credentials cleared")
    except Exception as e:
        logging.error(f"Failed to clear credentials: {type(e).__name__} - {str(e)}")
        messagebox.showerror("Error", f"無法清除憑證: {str(e)}")


# GUI 初始化
root = tk.Tk()
root.title("Phrase job editor (multi-projects)")
root.geometry("500x700")

frame_login = tk.Frame(root)
frame_login.pack(pady=10)
tk.Label(frame_login, text="帳號:").grid(row=0, column=0, padx=5)
entry_username = tk.Entry(frame_login)
entry_username.grid(row=0, column=1, padx=5)
tk.Label(frame_login, text="密碼:").grid(row=1, column=0, padx=5)
entry_password = tk.Entry(frame_login, show="*")
entry_password.grid(row=1, column=1, padx=5)
button_login = tk.Button(frame_login, text="登入", command=login)
button_login.grid(row=2, column=0, padx=5)
button_clear = tk.Button(frame_login, text="清除憑證", command=clear_credentials)
button_clear.grid(row=2, column=1, padx=5)
label_login_status = tk.Label(frame_login, text="未登入", foreground="red")
label_login_status.grid(row=3, column=0, columnspan=2)

frame_project = tk.Frame(root)
frame_project.pack(pady=10)
tk.Label(frame_project, text="專案名稱 (每行一個):").grid(row=0, column=0, padx=5, sticky="nw")
entry_project = tk.Text(frame_project, height=5, width=30, state="disabled")
entry_project.grid(row=0, column=1, padx=5)
tk.Label(frame_project, text="客戶名稱:").grid(row=1, column=0, padx=5)
entry_client = tk.Entry(frame_project, width=30, state="normal")
entry_client.grid(row=1, column=1, padx=5)
button_search = tk.Button(frame_project, text="查詢專案", command=search_project, state="disabled")
button_search.grid(row=0, column=2, padx=5, sticky="n")
label_project_uid = tk.Label(frame_project, text="Project UID: 未選擇")
label_project_uid.grid(row=2, column=0, columnspan=3, pady=5)

frame_workflow = tk.Frame(root)
frame_workflow.pack(pady=10)
tk.Label(frame_workflow, text="工作流程名稱:").grid(row=0, column=0, padx=5)

# 在預設選項中加入 "No Workflow"
all_values = [
    "No Workflow",
    "Translation", "Revision", "Client review", "Revision2",
    "Client Review2", "MT_pre", "Trans_pre", "Feedback",
    "MTPE", "Feedback2", "Client Review3"
]


def on_keyrelease(event):
    """當輸入文字時自動過濾選項"""
    value = event.widget.get().lower()
    if value == '':
        event.widget['values'] = all_values
    else:
        filtered = [item for item in all_values if value in item.lower()]
        event.widget['values'] = filtered


combo_workflow = ttk.Combobox(frame_workflow, values=all_values, state="normal")
combo_workflow.grid(row=0, column=1, padx=5)
combo_workflow.bind('<KeyRelease>', on_keyrelease)
tk.Label(frame_workflow, text="狀態:").grid(row=0, column=2, padx=5)
combo_status = ttk.Combobox(frame_workflow,
                            values=["NEW", "ACCEPTED", "DECLINED", "REJECTED", "DELIVERED", "EMAILED", "COMPLETED",
                                    "CANCELLED"], state="disabled")
combo_status.grid(row=0, column=3, padx=5)

# 新增語言選擇區域
frame_language = tk.Frame(root)
frame_language.pack(pady=10)
button_select_langs = tk.Button(frame_language, text="選擇目標語言", command=select_target_languages, state="normal")
button_select_langs.grid(row=0, column=0, padx=5)
label_selected_langs = tk.Label(frame_language, text="尚未選擇語言", foreground="blue")
label_selected_langs.grid(row=0, column=1, padx=5)

# 新增下載模式選擇區域
frame_download_mode = tk.Frame(root)
frame_download_mode.pack(pady=5)
tk.Label(frame_download_mode, text="下載模式:").grid(row=0, column=0, padx=5)
download_mode_var = tk.StringVar(value="合併下載")
radio_merge = tk.Radiobutton(frame_download_mode, text="合併下載 (專案名稱_語系)", variable=download_mode_var,
                             value="合併下載")
radio_merge.grid(row=0, column=1, padx=5)
radio_separate = tk.Radiobutton(frame_download_mode, text="單獨下載 (檔名_語系)", variable=download_mode_var,
                                value="單獨下載")
radio_separate.grid(row=0, column=2, padx=5)

frame_buttons = tk.Frame(root)
frame_buttons.pack(pady=10)
button_show_jobs = tk.Button(frame_buttons, text="顯示任務", command=show_jobs, state="disabled")
button_show_jobs.grid(row=0, column=0, padx=5)
button_update_status = tk.Button(frame_buttons, text="更新所有任務狀態", command=update_all_jobs_status,
                                 state="disabled")
button_update_status.grid(row=0, column=1, padx=5)
button_download_bilingual = tk.Button(frame_buttons, text="下載雙語檔案", command=download_bilingual_files_by_language,
                                      state="disabled")
button_download_bilingual.grid(row=0, column=2, padx=5)

frame_jobs = tk.Frame(root)
frame_jobs.pack(pady=10, fill=tk.BOTH, expand=True)
text_jobs = tk.Text(frame_jobs, height=50, width=60)
text_jobs.pack()


def initialize_app():
    credentials = load_credentials()
    if credentials:
        entry_username.insert(0, credentials["username"])
        if credentials.get("token") and is_token_valid(credentials.get("expires")):
            global API_TOKEN, HEADERS
            API_TOKEN = credentials["token"]
            HEADERS = {
                "Authorization": f"ApiToken {API_TOKEN}",
                "Content-Type": "application/json"
            }
            entry_project.config(state="normal")
            combo_workflow.config(state="normal")
            combo_status.config(state="normal")
            button_search.config(state="normal")
            button_show_jobs.config(state="normal")
            button_update_status.config(state="normal")
            button_download_bilingual.config(state="normal")
            label_login_status.config(text="已使用儲存的 Token", foreground="green")
            logging.info(f"Using stored token for user: {credentials['username']}")
        else:
            label_login_status.config(text="Token 無效或已過期,請登入", foreground="red")
            logging.info("Stored token invalid or expired, waiting for user login")


root.after(0, initialize_app)
root.mainloop()