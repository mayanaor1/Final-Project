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


def OpenGPT():
    # Initialize Chrome WebDriver options
    op = webdriver.ChromeOptions()
    op.add_argument(f"user-agent={UserAgent.random}")
    op.add_argument("user-data-dir=./")
    op.add_experimental_option("detach", True)
    op.add_experimental_option("excludeSwitches", ["enable-logging"])

    try:
        driver = uc.Chrome(chrome_options=op)
    except Exception as e:
        print("Error initializing WebDriver:", e)
        return None

    try:
        driver.get('https://chatgpt.com/')
        driver.maximize_window()

        # Wait until the "Continue" button is clickable on click on it
        continue_button = WebDriverWait(driver, 50).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "btn-primary"))
        )
        continue_button.click()

        # Wait for the username field to be visible, and then enter your username
        username_field = WebDriverWait(driver, 50).until(
            EC.visibility_of_element_located((By.ID, 'email-input'))
        )
        # username_field.send_keys('maya.naor@live.biu.ac.il')
        username_field.send_keys('mayathequeen123@gmail.com')
        # username_field.send_keys('naortalya@gmail.com')

        continue_button = driver.find_element(By.CLASS_NAME, "continue-btn")
        continue_button.click()

        # Wait for the password field to be visible, and then enter your password
        password_field = WebDriverWait(driver, 50).until(
            EC.visibility_of_element_located((By.ID, 'password'))
        )
        password_field.send_keys("Mn0546317685")
        submit_button = driver.find_element(By.CLASS_NAME, "_button-login-password")
        submit_button.click()

        # Pause execution for 5 seconds (to allow time for page loading)
        sleep(50)

        return driver
    except (TimeoutException, NoSuchElementException) as e:
        print("Error during login:", e)
        driver.quit()
        return None


def ChatGPT(question):
    try:
        # Wait until the "new chat" button is clickable on the page, and click on it
        sleep(1)
        new_button = WebDriverWait(driver, 50).until(
            EC.element_to_be_clickable((By.CLASS_NAME, 'text-token-text-primary'))
        )
        new_button.click()

        # Pause for 2 seconds to ensure the new chat interface is fully loaded
        sleep(2)

        # Locate the chat input field and wait until it is visible
        chat_input = WebDriverWait(driver, 50).until(
            EC.visibility_of_element_located((By.ID, 'prompt-textarea'))
        )
        # Enter the question into the chat input field
        chat_input.send_keys(question)
        print("Sent question:", question)
        # Simulate pressing the Enter key to send the question
        chat_input.send_keys(Keys.RETURN)

        WebDriverWait(driver, 30).until(
            lambda driver: any(
                text in driver.find_element(By.CLASS_NAME, 'agent-turn').text.lower() for text in ['bye', 'ביי'])
        )

        print("Waiting for response")
        response = ""
        # Wait until ChatGPT's response is visible on the page
        response = WebDriverWait(driver, 30).until(
            EC.visibility_of_element_located((By.CLASS_NAME, 'agent-turn'))
        )
        print("Received response")

        index = response.text.lower().find('bye')

        # If "bye" is found in the answer
        if index != -1:
            # Remove "bye" and everything after it
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


def SaveToFile(number):
    try:
        # Iterate over the rows in the first column of the worksheet, starting from the second row
        for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=1):
            # Retrieve the cell containing the question
            question_cell = row[0]
            # Extract the question text from the cell
            question = question_cell.value
            if question is None:
                return

            # Obtain the answer by sending the question to the ChatGPT function
            answer = ChatGPT(question)

            if answer:
                print(f"Answer: {answer}\n")

                # Write the answer to the cell next to the question cell (in column B)
                answer_cell = question_cell.offset(column=number)
                answer_cell.value = answer
                workbook.save('hebrew_profile6\Profile6Answers.xlsx')
                print("Saved answer to Excel")
    except Exception as e:
        print("Error saving to file:", e)


if __name__ == "__main__":
    # Initialize the WebDriver and open the chatbot interface
    driver = OpenGPT()

    if driver:
        # Define the path to the Excel file containing the questions
        excel_file_path = 'hebrew_profile6\profile6.xlsx'

        # Load the workbook from the specified Excel file path
        try:
            workbook = openpyxl.load_workbook(excel_file_path)
            sheet = workbook.active

            for i in range(1, 31):
                SaveToFile(i)

            print("answers saved to file successfully")
        except Exception as e:
            print("Error loading or processing Excel file:", e)

        # Close the browser window and end the WebDriver session
        driver.quit()
