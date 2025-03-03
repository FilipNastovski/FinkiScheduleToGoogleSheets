from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd
import time
import gspread
from google.oauth2 import service_account
import re


class ScheduleScraper:
    def __init__(self, url, credentials_path):
        """
        Initialize the scraper with the URL of the schedule page and path to Google API credentials
        """
        self.url = url
        self.driver = webdriver.Chrome()  # You may need to specify the path to chromedriver
        self.credentials_path = credentials_path
        self.schedule_data = {}

    def open_timetable(self):
        """
        Open the timetable page
        """
        try:
            self.driver.get(self.url)
            # Wait for the page to load
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            return True
        except Exception as e:
            print(f"Failed to open timetable: {e}")
            return False

    def select_class(self, class_name):
        """
        Click on Classes button and select the specified class
        """
        try:
            # Click the Classes button using its title attribute
            classes_button = self.driver.find_element(By.XPATH, "//div[@title='Classes']")
            classes_button.click()

            # Wait for the dropdown to appear
            time.sleep(1)

            # Find and click on your specific class
            class_element = self.driver.find_element(By.XPATH, f"//div[contains(text(), '{class_name}')]")
            class_element.click()

            # Wait for the timetable to update
            time.sleep(2)

            return True
        except Exception as e:
            print(f"Failed to select class: {e}")
            return False

    def select_subject(self, subject_name):
        """
        Click on Subjects button and select the specified subject
        """
        try:
            # Click the Subjects button using its title attribute
            subjects_button = self.driver.find_element(By.XPATH, "//div[@title='Subjects']")
            subjects_button.click()

            # Wait for the dropdown to appear
            time.sleep(1)

            # Find and click on the specified subject
            subject_element = self.driver.find_element(By.XPATH, f"//div[contains(text(), '{subject_name}')]")
            subject_element.click()

            # Wait for the timetable to update
            time.sleep(2)

            return True
        except Exception as e:
            print(f"Failed to select subject: {e}")
            return False

    def extract_schedule_data(self):
        """
        Extract data from the currently displayed SVG schedule
        """
        try:
            # Get the SVG element
            svg_element = self.driver.find_element(By.TAG_NAME, "svg")

            # Extract all text elements from the SVG
            text_elements = svg_element.find_elements(By.TAG_NAME, "text")

            # Extract data from SVG
            days = []
            times = []
            classes = []

            # First, identify days (should be the row headers in the timetable)
            day_names = ["Понеделник", "Вторник", "Среда", "Четврток", "Петок"]
            for element in text_elements:
                text = element.text
                if text in day_names:
                    position = element.get_attribute("x")
                    days.append({"name": text, "position": float(position)})

            # Sort days by position
            days.sort(key=lambda x: x["position"])

            # Next, identify time slots
            time_pattern = re.compile(r'\d+:\d+ - \d+:\d+')
            for element in text_elements:
                text = element.text
                if time_pattern.match(text):
                    position = element.get_attribute("y")
                    times.append({"time": text.split(" - ")[0], "position": float(position)})

            # Sort times by position
            times.sort(key=lambda x: x["position"])

            # Finally, extract class information
            class_elements = svg_element.find_elements(By.TAG_NAME, "rect")

            for element in class_elements:
                try:
                    # Skip transparent rectangles (these are usually just grid elements)
                    fill = element.get_attribute("fill")
                    if fill == "transparent" or not fill:
                        continue

                    # Get the title element which contains class information
                    title = element.find_element(By.TAG_NAME, "title")
                    class_info = title.get_attribute("innerHTML")

                    # Get position information
                    x = float(element.get_attribute("x"))
                    y = float(element.get_attribute("y"))
                    width = float(element.get_attribute("width"))
                    height = float(element.get_attribute("height"))

                    # Determine day based on x position
                    day = next((d["name"] for d in days if abs(d["position"] - (x + width / 2)) < 300), None)

                    # Determine time based on y position
                    time_slot = next((t["time"] for t in times if abs(t["position"] - y) < 100), None)

                    if day and time_slot and class_info:
                        # Parse class info into components (subject, professor, location)
                        info_parts = class_info.split("\n")
                        subject = info_parts[0] if len(info_parts) > 0 else ""
                        professor = info_parts[1] if len(info_parts) > 1 else ""
                        location = info_parts[2] if len(info_parts) > 2 else ""

                        classes.append({
                            "day": day,
                            "time": time_slot,
                            "subject": subject,
                            "professor": professor,
                            "location": location
                        })
                except Exception as e:
                    # Skip elements that don't have the expected structure
                    continue

            return classes
        except Exception as e:
            print(f"Failed to extract schedule data: {e}")
            return []

    def extract_class_schedule(self, class_name, subject_filter=None):
        """
        Select class and extract schedule, optionally filtering for specific subjects
        """
        if self.select_class(class_name):
            all_classes = self.extract_schedule_data()

            # Filter for specific subjects if needed
            if subject_filter:
                filtered_classes = [c for c in all_classes if
                                    any(subject in c["subject"] for subject in subject_filter)]
                return filtered_classes
            else:
                return all_classes
        return []

    def extract_subject_schedule(self, subject_names):
        """
        Extract schedule for specific subjects
        """
        all_subject_data = []

        for subject in subject_names:
            if self.select_subject(subject):
                subject_data = self.extract_schedule_data()
                all_subject_data.extend(subject_data)

        return all_subject_data

    def create_timetable_dataframe(self, schedule_data):
        """
        Create a DataFrame in the format requested
        """
        # Create empty DataFrame with time slots as index and days as columns
        days = ["Понеделник", "Вторник", "Среда", "Четврток", "Петок"]
        times = [
            "8:00", "9:00", "10:00", "11:00", "12:00",
            "13:00", "14:00", "15:00", "16:00", "17:00",
            "18:00", "19:00", "20:00"
        ]

        # Create MultiIndex DataFrame
        df = pd.DataFrame(index=times, columns=days)
        df.index.name = ""

        # Fill in the schedule data
        for entry in schedule_data:
            day = entry["day"]
            time = entry["time"]
            subject = entry["subject"]
            location = entry["location"]

            # Format: Subject (Location)
            if day in days and time in times:
                df.at[time, day] = f"{subject} {location}"

        # Reset index to get the time column
        df_reset = df.reset_index()

        # Add empty row at the beginning
        empty_row = pd.DataFrame([[""] * len(df_reset.columns)], columns=df_reset.columns)
        df_final = pd.concat([empty_row, df_reset], ignore_index=True)

        return df_final

    def save_to_google_sheets(self, data, sheet_name):
        """
        Save the collected schedule data to Google Sheets
        """
        try:
            # Set up credentials for Google Sheets API
            scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
            credentials = service_account.Credentials.from_service_account_file('service-account.json')
            gc = gspread.authorize(credentials)

            # Open or create the spreadsheet
            try:
                spreadsheet = gc.open(sheet_name)
            except:
                spreadsheet = gc.create(sheet_name)

            # Create or clear the worksheet
            try:
                worksheet = spreadsheet.worksheet("Schedule")
                worksheet.clear()
            except:
                worksheet = spreadsheet.add_worksheet(title="Schedule", rows=100, cols=20)

            # Convert the DataFrame to a list for gspread
            data_list = [data.columns.tolist()] + data.values.tolist()

            # Update the worksheet
            worksheet.update(data_list)

            print(f"Schedule data saved to Google Sheets: {sheet_name}")
            return True
        except Exception as e:
            print(f"Failed to save to Google Sheets: {e}")
            return False

    def close(self):
        """
        Close the browser
        """
        self.driver.quit()


def main():
    # Configuration
    url = "https://finki.edupage.org/timetable/"
    credentials_path = "google-credentials.json"  # Path to your Google API credentials

    # Subjects to extract from class view
    class_name = "3г-SEIS18"
    class_subjects = [
        "Интегрирани системи (п)",
        "Интегрирани системи (ав)",
        "Софтверски квалитет и тестирање"
    ]

    # Additional subjects to extract from subject view
    additional_subjects = [
        "Мултимедиски системи",
        "Оперативни системи",
        "Дизајн на интеракцијата човек-компјутер"
    ]

    # Create and run the scraper
    scraper = ScheduleScraper(url, credentials_path)

    # Open the timetable
    if scraper.open_timetable():
        # Extract class schedule data
        class_data = scraper.extract_class_schedule(class_name, class_subjects)

        # Extract additional subject data
        subject_data = scraper.extract_subject_schedule(additional_subjects)

        # Combine the data
        all_schedule_data = class_data + subject_data

        # Create timetable DataFrame
        timetable_df = scraper.create_timetable_dataframe(all_schedule_data)

        # Save to Google Sheets
        scraper.save_to_google_sheets(timetable_df, "FINKI Schedule")

    # Close the browser
    scraper.close()


if __name__ == "__main__":
    main()
