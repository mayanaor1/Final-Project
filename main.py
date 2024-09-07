import tkinter as tk
from tkinter import filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import undetected_chromedriver as uc
from fake_useragent import UserAgent
from time import sleep
import openpyxl
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from docx import Document
import os


def OpenGPT(username, password):
    """
    Initialize and open a Chrome WebDriver session for ChatGPT.

    This function sets up the Chrome WebDriver,
    navigates to the ChatGPT website, and logs in using the provided credentials.

    Args:
    username (str): The username for ChatGPT login.
    password (str): The password for ChatGPT login.

    Returns:
    WebDriver or None: The initialized WebDriver if successful, None otherwise.
    """
    # Initialize Chrome WebDriver options
    op = webdriver.ChromeOptions()
    op.add_argument(f"user-agent={UserAgent.random}")  # Use a random user agent
    op.add_argument("user-data-dir=./")  # Set user data directory
    op.add_experimental_option("detach", True)  # Keep browser open after script finishes
    op.add_experimental_option("excludeSwitches", ["enable-logging"])  # Disable logging

    try:
        driver = uc.Chrome(chrome_options=op)
    except Exception as e:
        print("Error initializing WebDriver:", e)
        return None

    try:
        driver.get('https://chatgpt.com/')  # Navigate to ChatGPT website
        driver.maximize_window()

        # Wait for and click the "Continue" button
        continue_button = WebDriverWait(driver, 50).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "btn-primary"))
        )
        continue_button.click()

        # Enter username
        username_field = WebDriverWait(driver, 50).until(
            EC.visibility_of_element_located((By.ID, 'email-input'))
        )
        username_field.send_keys(username)

        continue_button = driver.find_element(By.CLASS_NAME, "continue-btn")
        continue_button.click()

        # Enter password
        password_field = WebDriverWait(driver, 50).until(
            EC.visibility_of_element_located((By.ID, 'password'))
        )
        password_field.send_keys(password)
        submit_button = driver.find_element(By.CLASS_NAME, "_button-login-password")
        submit_button.click()

        sleep(50)  # Wait for page to load completely

        return driver
    except (TimeoutException, NoSuchElementException) as e:
        print("Error during login:", e)
        driver.quit()
        return None


def ChatGPT(question):
    """
    Send a question to ChatGPT and retrieve the response.

    This function interacts with the ChatGPT interface, sends the provided question,
    and waits for a response.

    Args:
    question (str): The question to be sent to ChatGPT.

    Returns:
    str or None: The response from ChatGPT if successful, None otherwise.
    """
    try:
        # Start a new chat
        sleep(2)
        new_button = WebDriverWait(driver, 50).until(
            EC.element_to_be_clickable((By.CLASS_NAME, 'text-token-text-primary'))
        )
        new_button.click()

        sleep(2)  # Wait for new chat interface to load

        # Enter and send the question
        chat_input = WebDriverWait(driver, 50).until(
            EC.visibility_of_element_located((By.ID, 'prompt-textarea'))
        )
        chat_input.send_keys(question)
        print("Sent question:", question)
        chat_input.send_keys(Keys.RETURN)

        # Wait for response
        WebDriverWait(driver, 30).until(
            lambda driver: any(
                text in driver.find_element(By.CLASS_NAME, 'agent-turn').text.lower() for text in ['bye', 'ביי'])
        )

        print("Waiting for response")
        response = WebDriverWait(driver, 30).until(
            EC.visibility_of_element_located((By.CLASS_NAME, 'agent-turn'))
        )
        print("Received response")

        # Remove "bye" and everything after it from the response
        index = response.text.lower().find('bye')
        if index != -1:
            response = response.text[:index]

        return response
    except (TimeoutException, NoSuchElementException) as e:
        error_message = driver.find_element(By.CLASS_NAME, "text-token-text-error")
        if error_message:
            print("You've reached our limit of messages per hour. Please try again later.")
            driver.quit()
            exit(1)
        else:
            print("Error in ChatGPT interaction:", e)
            driver.quit()
        return None


def SaveToFile(iterations):
    """
    Save ChatGPT responses to an Excel file.

    This function reads questions from an Excel file, sends them to ChatGPT,
    and saves the responses in a new Excel file.

    Args:
    iterations (int): The number of times to repeat the process for each question.

    Returns:
    str or None: The path to the new Excel file if successful, None otherwise.
    """
    try:
        # Create a new workbook for storing answers
        answer_wb = openpyxl.Workbook()
        answer_ws = answer_wb.active
        answer_ws.title = "Answers"

        # Write headers
        headers = ["Question"] + [f"Answer {i + 1}" for i in range(iterations)]
        answer_ws.append(headers)

        # Get all questions from the Excel file
        questions = [row[0].value for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=1)]

        for i in range(iterations):
            print(f"Starting iteration {i + 1}/{iterations}...")

            for question in questions:
                if not question:
                    continue

                row_idx = questions.index(question) + 2

                # Write question in first iteration
                if i == 0:
                    answer_ws.cell(row=row_idx, column=1, value=question)

                # Get answer from ChatGPT
                answer = ChatGPT(question)

                # Write answer
                answer_ws.cell(row=row_idx, column=i + 2, value=answer if answer else "No response")

            print(f"Iteration {i + 1} completed.")

        # Save answers to a new Excel file
        base_name = os.path.basename(excel_file_path)
        name, _ = os.path.splitext(base_name)
        new_excel_path = os.path.join(os.path.dirname(excel_file_path), f'{name}_Answers.xlsx')
        answer_wb.save(new_excel_path)
        print("Saved answers to new Excel file")

        return new_excel_path

    except Exception as e:
        print("Error saving to file:", e)
        return None


def SaveDocx(excel_path):
    """
    Convert the Excel file with ChatGPT responses to a Word document.

    This function reads the Excel file created by SaveToFile and creates a
    formatted Word document with questions and answers.

    Args:
    excel_path (str): The path to the Excel file containing ChatGPT responses.

    Returns:
    str or None: The path to the new Word document if successful, None otherwise.
    """
    try:
        # Load the Excel workbook
        wb = openpyxl.load_workbook(excel_path)
        sheet = wb.active

        # Create a new Word document
        doc = Document()

        # Process each row in the Excel file
        for row in sheet.iter_rows(min_row=2, values_only=True):
            cell1, *rest_of_cells = row

            # Add formatted question
            title1 = doc.add_paragraph("Question:")
            title1.runs[0].bold = True
            title1.runs[0].underline = True

            text_from_cell = str(cell1)
            question_mark_index = text_from_cell.find('?')
            modified_text = text_from_cell[:question_mark_index + 1]
            doc.add_paragraph(modified_text)

            # Add formatted answers
            title2 = doc.add_paragraph("Answers:")
            title2.runs[0].bold = True
            title2.runs[0].underline = True

            for cell_value in rest_of_cells:
                doc.add_paragraph(str(cell_value))

            doc.add_page_break()

        # Save the Word document
        docx_path = os.path.join(os.path.dirname(excel_file_path),
                                 f'{os.path.splitext(os.path.basename(excel_file_path))[0]}_Answers.docx')
        doc.save(docx_path)
        print("Saved document successfully")

        return docx_path

    except Exception as e:
        print("Error saving document:", e)
        return None


def run_script():
    """
    Main function to run the ChatGPT automation script.

    This function is triggered when the user clicks the "Start" button in the GUI.
    It orchestrates the entire process of interacting with ChatGPT, saving responses,
    and creating the final Word document.
    """
    username = username_entry.get()
    password = password_entry.get()
    iterations = int(iterations_spinbox.get())

    if not username or not password:
        messagebox.showerror("Input Error", "Please enter both username and password.")
        return

    global excel_file_path
    excel_file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
    )

    if not excel_file_path:
        messagebox.showerror("File Error", "Please select an Excel file.")
        return

    global driver
    driver = OpenGPT(username, password)

    if driver:
        try:
            global workbook, sheet
            workbook = openpyxl.load_workbook(excel_file_path)
            sheet = workbook.active

            new_excel_path = SaveToFile(iterations)

            if new_excel_path:
                SaveDocx(new_excel_path)

            driver.quit()

            # Delete the temporary Excel file
            if new_excel_path and os.path.exists(new_excel_path):
                os.remove(new_excel_path)
                print(f"Deleted the new Excel file: {new_excel_path}")
                messagebox.showinfo("Success", "Answers saved to Word document successfully")

        except Exception as e:
            driver.quit()
            messagebox.showerror("Error", f"Error processing file: {e}")


# Create the main window
root = tk.Tk()
root.title("ChatGPT Automation")

# Create and place widgets
tk.Label(root, text="ChatGPT username:").grid(row=0, column=0, padx=10, pady=10)
username_entry = tk.Entry(root, width=30)
username_entry.grid(row=0, column=1, padx=10, pady=10)

tk.Label(root, text="ChatGPT password:").grid(row=1, column=0, padx=10, pady=10)
password_entry = tk.Entry(root, width=30, show='*')
password_entry.grid(row=1, column=1, padx=10, pady=10)

tk.Label(root, text="Iterations:").grid(row=2, column=0, padx=10, pady=10)
iterations_spinbox = tk.Spinbox(root, from_=1, to=100, width=5)
iterations_spinbox.grid(row=2, column=1, padx=10, pady=10)

start_button = tk.Button(root, text="Start", command=run_script)
start_button.grid(row=3, column=0, columnspan=2, padx=10, pady=10)

# Run the GUI event loop
root.mainloop()
