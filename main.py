from playwright.sync_api import sync_playwright
import pandas as pd
import time
import logging
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import os

# --- Setup Logging ---
logging.basicConfig(filename='epfo_scraper.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

def login(page, url, username, password):
    """Logs into the EPFO portal with MANUAL CAPTCHA entry."""
    page.goto(url)
    try:
        page.fill("#username1", username)
        print("Username field located successfully")
        page.fill("#password", password)
        print("Password field located successfully")
        page.fill("#captcha", "")
        print("Found the captcha")
        page.click('//*[@id="AuthenticationForm"]/div[4]/div[1]/button')
        print("Login button located and clickable")

        # --- MANUAL CAPTCHA Entry ---
        messagebox.showinfo("CAPTCHA Required", "Please solve the CAPTCHA in the browser window that just opened.\nThen, enter the CAPTCHA in the GUI and click 'Submit CAPTCHA' to continue.")
        
        # Wait for the user to enter the CAPTCHA in the GUI and click the submit button
        captcha_entry.wait_variable(captcha_var)
        manual_captcha = captcha_entry.get()
        
        page.fill("#captcha", manual_captcha) # Enter manually inputted captcha

        page.click('//*[@id="AuthenticationForm"]/div[4]/div[1]/button')

        # --- Wait for Successful Login (Adapt if Necessary) ---
        page.wait_for_selector("//div[@class='container-fluid']", timeout=120000) # Wait for an element that appears ONLY after login
        logging.info("Login successful.")
        return True

    except Exception as e:
        logging.error(f"Login failed: {e}")
        messagebox.showerror("Login Error", f"Login failed: {e}")
        return False

def extract_data_for_uan(page, uan):
    """Extracts data for a single UAN after navigating to the member search."""
    try:
        # --- Navigate to Member Search ---
        search_input = page.locator('xpath=/html/body/div/div/form/div[1]/div[2]/div/div[1]/div[1]/div/div[2]/input')
        search_button = page.locator('xpath=/html/body/div/div/form/div[1]/div[2]/div/div[2]/div/input')

        search_input.fill(uan)
        search_button.click()

        # --- Wait for Search Results ---
        page.wait_for_selector("//div[@class='container-fluid']", timeout=120000)
        # --- Extract Data Using Labels ---
        name_element = page.locator('xpath=/html/body/div/div/div/div/div/div[2]/table/tbody/tr[2]/td[2]')
        name = name_element.inner_text().strip() # Use .inner_text() and locator

        joining_date_element = page.locator('xpath=/html/body/div/div/div/div/div/div[2]/table/tbody/tr[3]/td[4]')
        joining_date = joining_date_element.inner_text().strip() # Use .inner_text() and locator

        exit_date_element = page.locator('xpath=/html/body/div/div/div/div/div/div[2]/table/tbody/tr[4]/td[2]')
        exit_date = exit_date_element.inner_text().strip() # Use .inner_text() and locator

        logging.info(f"Data extracted for UAN: {uan}")
        return {
            "UAN": uan,
            "Name": name,
            "Joining Date": joining_date,
            "Exit Date": exit_date,
        }

    except Exception as e:
        logging.error(f"Could not extract all data for UAN {uan}: {e}")
        return None

def run_extraction():
    """Gets input from GUI, runs the extraction, and saves to Excel."""
    username = username_entry.get()
    password = password_entry.get()
    uans_text = uans_entry.get("1.0", tk.END)  # Get text from Text widget
    output_file = output_file_entry.get()

    # Basic input validation
    if not username or not password:
        messagebox.showerror("Error", "Please enter both username and password.")
        return
    if not uans_text.strip():
        messagebox.showerror("Error", "Please enter at least one UAN.")
        return
    if not output_file:
        messagebox.showerror("Error", "Please enter an output file name.")
        return
    uans = [u.strip() for u in uans_text.split(',') if u.strip()]
    print(uans)

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=False)
        page = browser.new_page()

        # --- Login ---
        url = "https://unifiedportal-emp.epfindia.gov.in/epfo/"
        if not login(page, url, username, password):
            print("Unable to login")
            browser.close()
            return

        # --- Navigate to member tab after login ---
        try:
            member_tab = page.locator('xpath=//*[@id="menu"]/li[2]/a') # Explicit XPath
            member_tab.click()

            member_info_option = page.locator('xpath=//*[@id="menu"]/li[2]/ul/li[1]/a') # Explicit XPath
            member_info_option.click()
            # page.click('//*[@id="menu"]/li[2]/a') # Member Tab Button
            # page.click('//*[@id="menu"]/li[2]/ul/li[1]/a')# Member Profile Option

        except Exception as e:
            messagebox.showerror("Navigation Error", f"Could not navigate to 'Member Information' page: {e}")
            logging.error(f"Navigation to Member Information page failed: {e}")
            browser.close()
            return

        # --- Extract Data for Each UAN ---
        all_uan_data = []
        for uan in uans:
            print("extracting...")
            uan_data = extract_data_for_uan(page, uan)
            if uan_data:
                all_uan_data.append(uan_data)
            # Add a small delay to avoid overwhelming the server
            time.sleep(2)

        # --- Save Data to Excel ---
        if all_uan_data:
            try:
                df = pd.DataFrame(all_uan_data)
                if not output_file.endswith(".xlsx"):
                    output_file += ".xlsx"  # Ensure .xlsx extension
                # Remove the existing file if it exists
                if os.path.exists(output_file):
                    os.remove(output_file)
                df.to_excel(output_file, index=False)
                messagebox.showinfo("Success", f"Data extracted and saved to {output_file}")
                logging.info(f"Data saved to {output_file}")

            except Exception as e:
                messagebox.showerror("Excel Error", f"Failed to save data to Excel: {e}")
                logging.error(f"Failed to save data to Excel: {e}")
        else:
            messagebox.showinfo("Info", "No data was extracted.")
            logging.warning("No data was extracted.")

        browser.close()

def browse_file():
    """Opens a file dialog to select the output file."""
    filename = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])
    if filename:
        output_file_entry.delete(0, tk.END)
        output_file_entry.insert(0, filename)

def submit_captcha():
    """Callback function for the CAPTCHA submit button."""
    captcha_var.set(1)

# --- GUI Setup ---
root = tk.Tk()
root.title("EPFO Data Extractor")

# --- Labels and Entry Fields ---
ttk.Label(root, text="Username:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
username_entry = ttk.Entry(root, width=40)
username_entry.grid(row=0, column=1, padx=5, pady=5)

ttk.Label(root, text="Password:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
password_entry = ttk.Entry(root, width=40, show="*")  # Show asterisks for password
password_entry.grid(row=1, column=1, padx=5, pady=5)

ttk.Label(root, text="UANs (comma-separated):").grid(row=2, column=0, padx=5, pady=5, sticky="w")
uans_entry = scrolledtext.ScrolledText(root, width=40, height=5)  # Use ScrolledText for multiple UANs
uans_entry.grid(row=2, column=1, padx=5, pady=5)

ttk.Label(root, text="Output File:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
output_file_entry = ttk.Entry(root, width=30)
output_file_entry.grid(row=3, column=1, padx=5, pady=5, sticky="we")
output_file_entry.insert(0, "epfo_data.xlsx") #Default file

browse_button = ttk.Button(root, text="Browse", command=browse_file)
browse_button.grid(row=3, column=2, padx=5, pady=5)

# --- CAPTCHA Entry ---
ttk.Label(root, text="CAPTCHA:").grid(row=4, column=0, padx=5, pady=5, sticky="w")
captcha_var = tk.IntVar()
captcha_entry = ttk.Entry(root, width=40)
captcha_entry.grid(row=4, column=1, padx=5, pady=5)

submit_captcha_button = ttk.Button(root, text="Submit CAPTCHA", command=submit_captcha)
submit_captcha_button.grid(row=4, column=2, padx=5, pady=5)

# --- Run Button ---
run_button = ttk.Button(root, text="Extract Data", command=run_extraction)
run_button.grid(row=5, column=1, padx=5, pady=10)

root.mainloop()