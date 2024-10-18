import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import re

# Set up Chrome options to suppress logging
chrome_options = Options()
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
chrome_options.add_experimental_option('w3c', False)

if getattr(sys, 'frozen', False):
    chromedriver_path = os.path.join(sys._MEIPASS, 'chromedriver.exe')
else:
    chromedriver_path = 'C:\\Users\\scott\\Desktop\\Python\\chromedriver.exe'

service = Service(chromedriver_path)
driver = webdriver.Chrome(service=service)

results = []  # Store search results globally

def is_blacklisted(url, blacklist):
    for word in blacklist:
        if word.lower() in url.lower():
            return True
    return False

def google_search(query, num_results, num_pages, start_page, blacklist):
    global results
    results = []  # Reset results for each new search
    try:
        driver.get('https://www.google.com')
        search_box = driver.find_element(By.NAME, 'q')
        search_box.send_keys(query)
        search_box.send_keys(Keys.RETURN)

        start_page = max(0, start_page - 1)
        for page in range(start_page, start_page + num_pages):
            if page > 0:
                driver.get(f'https://www.google.com/search?q={query}&start={page * 10}')

            WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '//h3')))
            time.sleep(2)  # Wait for the page to fully load

            search_results = driver.find_elements(By.XPATH, '//a[not(ancestor::div[contains(@aria-label, "Ads")])]/h3/ancestor::a')

            for result in search_results:
                if len(results) >= num_results:
                    break
                link = result.get_attribute('href')
                if link and not is_blacklisted(link, blacklist) and "google.com" not in link:
                    results.append(link)

            progress_bar["value"] += 1
            root.update_idletasks()

            if len(results) >= num_results:
                break

        return results
    except Exception as e:
        print(f"Error during Google search: {e}")
        return []

def extract_contact_info(elements):
    email_pattern = r'(?<![\w.-])([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})(?![\w.-])'
    phone_pattern = r'(?:Telephone|Toll Free|Fax)?\s*:?[-\s]?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}'

    emails = set()
    phones = set()

    for element in elements:
        text = element.text
        found_emails = re.findall(email_pattern, text)
        emails.update(found_emails)
        found_phones = re.findall(phone_pattern, text)
        phones.update(found_phones)

    phones = [re.sub(r'(?:Telephone|Toll Free|Fax)?\s*:?[-\s]?', '', phone).strip() for phone in phones]

    filtered_emails = [email for email in emails if not re.match(r'^[a-zA-Z0-9]{12,}$', email.split('@')[0])]

    return {"emails": set(filtered_emails), "phones": set(phones)}

def find_contact_links():
    try:
        links = driver.find_elements(By.XPATH, '//a[@href]')
        contact_links = set()
        keywords = ["contact", "about", "support", "help", "customer", "feedback"]
        for link in links:
            url = link.get_attribute('href')
            if any(keyword in url.lower() for keyword in keywords):
                contact_links.add(url)
        return contact_links
    except Exception as e:
        print(f"Error finding contact links: {e}")
        return set()

def scrape_website(url):
    try:
        driver.get(url)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
        visible_elements = driver.find_elements(By.XPATH, "//body//*")
        return extract_contact_info(visible_elements)
    except Exception as e:
        print(f"Error scraping {url}: {e}")
        return {}

def scrape_website_with_contact_pages(url):
    try:
        main_contact_info = scrape_website(url)
        contact_links = find_contact_links()

        all_contact_info = {
            "emails": set(main_contact_info["emails"]),
            "phones": set(main_contact_info["phones"]),
        }

        for contact_url in contact_links:
            print(f"Scraping contact page: {contact_url}")
            contact_info = scrape_website(contact_url)
            all_contact_info["emails"].update(contact_info.get("emails", set()))
            all_contact_info["phones"].update(contact_info.get("phones", set()))

        return all_contact_info
    except Exception as e:
        print(f"Error scraping {url}: {e}")
        return {}

def display_results(results, listbox):
    listbox.delete(0, tk.END)
    for i, result in enumerate(results):
        listbox.insert(tk.END, f"{i + 1}. {result}")

def search_and_display():
    query = search_query.get()
    num_results = int(num_results_entry.get())
    num_pages = int(num_pages_entry.get())
    start_page = int(start_page_entry.get())

    blacklist_input = blacklist_text.get("1.0", tk.END)
    blacklist = [word.strip() for word in blacklist_input.splitlines() if word.strip()]

    progress_bar["maximum"] = num_pages
    progress_bar["value"] = 0

    results = google_search(query, num_results, num_pages, start_page, blacklist)
    if not results:
        messagebox.showerror("Error", "No results found or an error occurred.")
        return

    display_results(results, result_listbox)
    progress_bar["maximum"] = len(results)
    progress_bar["value"] = 0

def scrape_selected():
    selected_indices = result_listbox.curselection()
    if not selected_indices:
        messagebox.showerror("Error", "Please select at least one result to scrape.")
        return

    result_text.delete(1.0, tk.END)
    no_contact_info_count = 0
    progress_bar["maximum"] = len(selected_indices)
    progress_bar["value"] = 0

    for index in selected_indices:
        url = results[index]
        contact_info = scrape_website_with_contact_pages(url)

        emails = contact_info.get('emails', set())
        phones = contact_info.get('phones', set())

        if emails or phones:
            result_text.insert(tk.END, f"Contact info from {url}:\n")
            if emails:
                result_text.insert(tk.END, f"Emails: {', '.join(emails)}\n")
            if phones:
                result_text.insert(tk.END, f"Phones: {', '.join(phones)}\n")
            result_text.insert(tk.END, "\n")
        else:
            no_contact_info_count += 1

        progress_bar["value"] += 1
        root.update_idletasks()

    if no_contact_info_count > 0:
        result_text.insert(tk.END, f"{no_contact_info_count} sites had no email or phone information.\n")

def export_contact_info():
    contact_info = result_text.get("1.0", tk.END).strip()
    if not contact_info:
        messagebox.showerror("Error", "No contact information to export.")
        return

    file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt"), ("CSV files", "*.csv"), ("All files", "*.*")])
    if file_path:
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(contact_info)
            messagebox.showinfo("Success", "Contact information exported successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save the file: {e}")

# Create the main window with a dark grey theme
root = tk.Tk()
root.title("Web Scraper GUI")
root.configure(bg="#2e2e2e")

# Apply ttk style for a more polished look
style = ttk.Style()
style.theme_use("clam")
style.configure("TFrame", background="#2e2e2e")
style.configure("TLabel", background="#2e2e2e", foreground="white")
style.configure("TButton", background="#444", foreground="white")
style.configure("TEntry", fieldbackground="#444", foreground="white", insertcolor="white")
style.configure("TListbox", background="#444", foreground="white", selectbackground="#666")
style.configure("TProgressbar", troughcolor="#888", background="Green")  # Change background to Green

# Create the left frame for inputs and buttons
left_frame = ttk.Frame(root)
left_frame.pack(side=tk.LEFT, padx=20, pady=20, anchor=tk.N)

# Center-align the widgets inside the left frame
left_frame.grid_columnconfigure(0, weight=1)

# Search query label and entry
search_label = ttk.Label(left_frame, text="Search Query:")
search_label.grid(row=0, column=0, pady=5, sticky="ew")

search_query = ttk.Entry(left_frame, width=50)
search_query.grid(row=1, column=0, pady=5, sticky="ew")

# Number of results label and entry
num_results_label = ttk.Label(left_frame, text="Number of results to retrieve:")
num_results_label.grid(row=2, column=0, pady=5, sticky="ew")

num_results_entry = ttk.Entry(left_frame, width=10)
num_results_entry.insert(0, "10")
num_results_entry.grid(row=3, column=0, pady=5, sticky="ew")

# Number of pages label and entry
num_pages_label = ttk.Label(left_frame, text="Number of Google pages to search:")
num_pages_label.grid(row=4, column=0, pady=5, sticky="ew")

num_pages_entry = ttk.Entry(left_frame, width=10)
num_pages_entry.insert(0, "1")
num_pages_entry.grid(row=5, column=0, pady=5, sticky="ew")

# Starting page label and entry
start_page_label = ttk.Label(left_frame, text="Starting Google page:")
start_page_label.grid(row=6, column=0, pady=5, sticky="ew")

start_page_entry = ttk.Entry(left_frame, width=10)
start_page_entry.insert(0, "1")
start_page_entry.grid(row=7, column=0, pady=5, sticky="ew")

# Blacklist label and entry
blacklist_label = ttk.Label(left_frame, text="Blacklist (one word per line):")
blacklist_label.grid(row=8, column=0, pady=5, sticky="ew")

blacklist_text = tk.Text(left_frame, width=50, height=5, bg="#444", fg="white")
blacklist_text.grid(row=9, column=0, pady=5, sticky="ew")

# Search button
search_button = ttk.Button(left_frame, text="Search", command=search_and_display)
search_button.grid(row=10, column=0, pady=10, sticky="ew")

# Create a spacer row
spacer_label = ttk.Label(left_frame, text="", background="#2e2e2e")
spacer_label.grid(row=11, column=0, pady=5)

# Scrape button
scrape_button = ttk.Button(left_frame, text="Scrape Selected", command=scrape_selected)
scrape_button.grid(row=12, column=0, pady=5, sticky="ew")

# Export button
export_button = ttk.Button(left_frame, text="Export Contact Info", command=export_contact_info)
export_button.grid(row=13, column=0, pady=5, sticky="ew")

# Progress bar
progress_bar = ttk.Progressbar(left_frame, orient="horizontal", length=200, mode="determinate")
progress_bar.grid(row=14, column=0, pady=10, sticky="ew")

# Create the right frame for displaying results
right_frame = ttk.Frame(root)
right_frame.pack(side=tk.RIGHT, padx=20, pady=20)

# Result list box
result_listbox = tk.Listbox(right_frame, selectmode=tk.MULTIPLE, width=70, height=15, bg="#444", fg="white", bd=0)
result_listbox.pack(pady=10)

# Result text area
result_text = scrolledtext.ScrolledText(right_frame, width=70, height=15, bg="#444", fg="white", bd=0, highlightthickness=0)
result_text.pack(pady=10)

# Start the GUI event loop
root.mainloop()

# Clean up the WebDriver when the GUI is closed
driver.quit()
