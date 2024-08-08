import re
import time
import warnings

import pandas as pd
from IPython.display import display
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def format_device_code(expected_value):
    if isinstance(expected_value, list):
        if len(expected_value) == 1:
            return expected_value[0]
        elif len(expected_value) > 1:
            return " or ".join(expected_value)
    return expected_value


def modify_expected_value(row):
    parameter = row["Parameters"]
    if parameter in ["Number of S/C", "Number of C/C"]:
        return "(Digit, Not Blank)"
    elif parameter == "Application":
        return "(Type of Application, Not Blank)"
    else:
        return "(Not Blank)" if not row["Should be Equal"] else row["Expected Value"]


# Apply styles to the DataFrame
def highlight_row(x):
    if x["Parameters"] != "Device Code":
        if (x["Actual Value"] != x["Expected Value"] and x["Should be Equal"]) or (
            not x["Should be Equal"] and x["Actual Value"] == ""
        ):
            return ["color: red; font-weight: bold"] * len(x)
    elif x["Parameters"] == "Device Code" and x["Actual Value"] not in x[
        "Expected Value"
    ].split(" or "):
        return ["color: red; font-weight: bold"] * len(x)
    return ["color: black; font-weight: normal"] * len(x)


class Utils:
    """Clase que contiene funciones y métodos útiles para la automatización."""

    def __init__(self, extra_time=0):
        self.extra_time_sec = int(extra_time)
        self.letter_identifier = None

    def wait_for_element(self, driver, locator, timeout=10):
        """Wait for element to be present on the page."""
        return WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located(locator)
        )

    def click_element(self, driver, locator):
        """Click on the element identified by the locator."""
        element = self.wait_for_element(driver, locator)
        element.click()

    def open_url(self, url, driver_type="edge"):
        """Open the given URL in a new browser window."""
        with warnings.catch_warnings():
            warnings.filterwarnings(
                "ignore", message="selenium.webdriver.common.selenium_manager"
            )
            if driver_type == "edge":
                driver = webdriver.Edge()
            elif driver_type == "firefox":
                driver = webdriver.Firefox()
            elif driver_type == "chrome":
                driver = webdriver.Chrome()
            else:
                raise ValueError(
                    "Invalid driver type. Supported types are 'chrome', 'firefox', and 'edge'."
                )
        driver.get(url)
        time.sleep(10 + self.extra_time_sec)
        self.wait_for_element(driver, (By.TAG_NAME, "body"), timeout=50)
        return driver

    def close_session(self, driver):
        if driver:
            driver.find_element(
                By.XPATH,
                f"//*[@id='main-view']/div/div[2]/div[2]/nav/div[1]/div[2]/button",
            ).click()
            try:
                time.sleep(3)
                driver.find_element(
                    By.XPATH,
                    f"//*[@id='globalNavigationSideNav']/div[2]/div/div[4]/div[2]/div/div[2]/button",
                ).click()
            except TimeoutException:
                time.sleep(2)
                driver.find_element(
                    By.XPATH,
                    f"//*[@id='globalNavigationSideNav']/div[2]/div/div[4]/div[2]/div/div[2]/button",
                ).click()
            except Exception as e:
                print("This is an Exception message. Unable to Close the browser.")
            time.sleep(4)
            driver.quit()

    def search_for_part(self, driver, part_number):
        """Search for a part number on the webpage."""
        search_element = self.wait_for_element(
            driver, (By.CSS_SELECTOR, "div.aw-search-globalSearchWidgetContainer")
        )
        search_element.click()
        time.sleep(1 + 0.25 * self.extra_time_sec)
        selector_element = self.wait_for_element(
            driver,
            (
                By.XPATH,
                "//*[@id='aw_navigation']//div[@class='aw-search-searchPreFilterPanel2']",
            ),
        )
        selector_element.click()
        self.wait_for_element(
            driver, (By.XPATH, "//div[@class='sw-cell-valName'][@title='Parts']")
        )
        time.sleep(1 + 0.25 * self.extra_time_sec)
        parts_element = self.wait_for_element(
            driver, (By.XPATH, "//div[@class='sw-cell-valName'][@title='Parts']")
        )
        parts_element.click()
        input_element = self.wait_for_element(
            driver,
            (
                By.XPATH,
                "//*[@id='aw_navigation']/div[2]/form/div[2]/div/div/div[1]/div/div[1]/div[2]/div/div/input",
            ),
        )
        input_element.send_keys(part_number)
        time.sleep(1 + 0.25 * self.extra_time_sec)
        search_button = self.wait_for_element(
            driver,
            (
                By.XPATH,
                "//*[@id='aw_navigation']/div[2]/form/div[2]/div/div/div[1]/div/div[1]/div[2]/div/div/div[2]",
            ),
        )
        search_button.click()

    def select_from_search(self, driver, part_number):
        """Select the part from the search results."""
        time.sleep(15 + self.extra_time_sec)
        li_elements = driver.find_elements(
            By.CSS_SELECTOR, "ul.aw-widgets-cellListWidget > li"
        )
        for i, li_element in enumerate(li_elements):
            time.sleep(1 + 0.25 * self.extra_time_sec)
            span_title_element = li_element.find_element(
                By.CSS_SELECTOR, "span.aw-widgets-cellListCellTitle"
            )
            title_text = span_title_element.get_attribute("title")
            part_num = title_text.split(",")[0].strip()
            if part_num.upper() == part_number.upper():
                time.sleep(1 + 0.25 * self.extra_time_sec)
                driver.find_element(
                    By.XPATH, f"//*[@id='main-view']//ul/li[{i+1}]/div/div[2]/div/div"
                ).click()
                return True
        print("\033[48;5;208mNo matching result for the Part Number\n\033[0m")
        return False

    def get_attr_from_part_search(self, driver):
        wait = WebDriverWait(driver, 20)
        title_id = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "span[data-locator='ID']"))
        ).get_attribute("value")
        title_id_value = wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "span[data-locator='Title ID Value']")
            )
        ).get_attribute("value")
        life_cycle_state = wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "span[data-locator='Lifecycle State']")
            )
        ).get_attribute("value")
        type_number = wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "span[data-locator='Type Number']")
            )
        ).get_attribute("value")
        customers = self.get_attribute(driver, "Saleable only to Customer(s)")
        if not type_number:
            type_number = ""
        return title_id, title_id_value, life_cycle_state, type_number.upper(), customers

    def expand_sections(self, driver):
        section_titles = ["Product", "Special Characteristics"]
        for section in section_titles:
            self.click_element(driver, (By.CSS_SELECTOR, f"div[title='{section}']"))

    def attributes(self, part_number, type_dropdown, letter_identifier) -> dict:
        match = re.match(r"K\d+", part_number)
        part = match.group(0) if match else None
        required_attributes = {
            "FLC-20": {
                "R": {
                    "Main Info": {
                        "equal_fields": {
                            "Article Type Value": "AF; Finished Part",
                            "Saleable Item Type Value": "IP; Product",
                            "Title ID Value": "Camera; 2687",
                            "Use Status": "Series",
                            "Type Number": type_dropdown.value,
                        },
                        "no_empty_fields": [],
                    },
                    "Product": {
                        "equal_fields": {
                            "Product Group": "29LL",
                            "Product Group Disp": "Forward Looking Camera",
                            "Device Code": ["21024", "50017"],
                            "Sales Channel Limitation": "OEM only",
                            "Price Book": "BXA (air products)",
                            "Basic Part Number": part,
                        },
                        "no_empty_fields": ["Saleable only to Customer(s)"],
                    },
                    "Special Characteristics": {
                        "equal_fields": {},
                        "no_empty_fields": ["Number of S/C", "Number of C/C"],
                    },
                    "Classification": {
                        "equal_fields": {},
                        "no_empty_fields": [
                            "Application",
                            "ECU Type",
                            "CAN Baud Rate",
                            "Camera Angle",
                            "Vehicle Connector",
                            "Video Connector",
                            "Camera Cable Length",
                            "Software Version",
                        ],
                    },
                },
                "X": {
                    "Main Info": {
                        "equal_fields": {
                            "Article Type Value": "AF; Finished Part",
                            "Saleable Item Type Value": "N; Not for external sale",
                            "Title ID Value": "Camera; 2687",
                            "Use Status": "Series",
                            "Type Number": type_dropdown.value,
                        },
                        "no_empty_fields": [],
                    },
                    "Product": {
                        "equal_fields": {
                            "Product Group": "29LL",
                            "Product Group Disp": "Forward Looking Camera",
                            "Device Code": ["21024", "50017"],
                            "Sales Channel Limitation": "Inter company only",
                            "Price Book": "BXA (air products)",
                            "Basic Part Number": part,
                        },
                        "no_empty_fields": [],
                    },
                    "Special Characteristics": {
                        "equal_fields": {},
                        "no_empty_fields": ["Number of S/C", "Number of C/C"],
                    },
                    "Classification": {
                        "equal_fields": {},
                        "no_empty_fields": [
                            "Application",
                            "ECU Type",
                            "CAN Baud Rate",
                            "Camera Angle",
                            "Vehicle Connector",
                            "Video Connector",
                            "Camera Cable Length",
                            "Software Version",
                        ],
                    },
                },
                "N": {
                    "Main Info": {
                        "equal_fields": {
                            "Article Type Value": "AF; Finished Part",
                            "Saleable Item Type Value": "IE; Exchange Product",
                            "Title ID Value": "Camera; 2687",
                            "Use Status": "Series",
                            "Type Number": type_dropdown.value,
                        },
                        "no_empty_fields": [],
                    },
                    "Product": {
                        "equal_fields": {
                            "Product Group": "29LL",
                            "Product Group Disp": "Forward Looking Camera",
                            "Device Code": ["21024", "50017"],
                            "Saleable only to Customer(s)": "",
                            "Sales Channel Limitation": "AM only",
                            "Price Book": "BXA (air products)",
                            "Basic Part Number": part,
                        },
                        "no_empty_fields": [],
                    },
                    "Special Characteristics": {
                        "equal_fields": {},
                        "no_empty_fields": ["Number of S/C", "Number of C/C"],
                    },
                    "Classification": {
                        "equal_fields": {},
                        "no_empty_fields": [
                            "Application",
                            "ECU Type",
                            "CAN Baud Rate",
                            "Camera Angle",
                            "Vehicle Connector",
                            "Video Connector",
                            "Camera Cable Length",
                            "Software Version",
                        ],
                    },
                },
                "SC": {
                    "Main Info": {
                        "equal_fields": {
                            "Article Type Value": "AF; Finished Part",
                            "Saleable Item Type Value": "IP; Product",
                            "Title ID Value": "Camera; 2687",
                            "Use Status": "Series",
                            "Type Number": type_dropdown.value,
                        },
                        "no_empty_fields": [],
                    },
                    "Product": {
                        "equal_fields": {
                            "Product Group": "29LL",
                            "Product Group Disp": "Forward Looking Camera",
                            "Device Code": ["21024", "50017"],
                            "Saleable only to Customer(s)": "All Customers",
                            "Sales Channel Limitation": "AM only",
                            "Price Book": "BXA (air products)",
                            "Basic Part Number": part,
                        },
                        "no_empty_fields": [],
                    },
                    "Special Characteristics": {
                        "equal_fields": {},
                        "no_empty_fields": ["Number of S/C", "Number of C/C"],
                    },
                    "Classification": {
                        "equal_fields": {},
                        "no_empty_fields": [
                            "Application",
                            "ECU Type",
                            "CAN Baud Rate",
                            "Camera Angle",
                            "Vehicle Connector",
                            "Video Connector",
                            "Camera Cable Length",
                            "Software Version",
                        ],
                    },
                },
            },
            "FLR-21": {
                "R": {
                    "Main Info": {
                        "equal_fields": {
                            "Article Type Value": "AF; Finished Part",
                            "Saleable Item Type Value": "IP; Product",
                            "Title ID Value": "Front Radar Assembly; 3753",
                            "Use Status": "Series",
                            "Type Number": type_dropdown.value,
                        },
                        "no_empty_fields": [],
                    },
                    "Product": {
                        "equal_fields": {
                            "Product Group": "29LA",
                            "Product Group Disp": "Forward Looking Radar",
                            "Device Code": ["4963", "50013", "50016"],
                            "Sales Channel Limitation": "OEM only",
                            "Price Book": "BXA (air products)",
                            "Basic Part Number": part,
                        },
                        "no_empty_fields": ["Saleable only to Customer(s)"],
                    },
                    "Special Characteristics": {
                        "equal_fields": {},
                        "no_empty_fields": ["Number of S/C", "Number of C/C"],
                    },
                    "Classification": {
                        "equal_fields": {},
                        "no_empty_fields": [
                            "Application",
                            "Product Type",
                            "CAN Baud Rate",
                            "Connector",
                            "Connector 2",
                            "Software Version",
                            "Maximum Operating Temperature",
                            "Minimum Operating Temperature",
                        ],
                    },
                },
                "X": {
                    "Main Info": {
                        "equal_fields": {
                            "Article Type Value": "AF; Finished Part",
                            "Saleable Item Type Value": "N; Not for external sale",
                            "Title ID Value": "Front Radar Assembly; 3753",
                            "Use Status": "Series",
                            "Type Number": type_dropdown.value,
                        },
                        "no_empty_fields": [],
                    },
                    "Product": {
                        "equal_fields": {
                            "Product Group": "29LA",
                            "Product Group Disp": "Forward Looking Radar",
                            "Device Code": ["4963", "50013", "50016"],
                            "Sales Channel Limitation": "Inter company only",
                            "Price Book": "BXA (air products)",
                            "Basic Part Number": part,
                        },
                        "no_empty_fields": [],
                    },
                    "Special Characteristics": {
                        "equal_fields": {},
                        "no_empty_fields": ["Number of S/C", "Number of C/C"],
                    },
                    "Classification": {
                        "equal_fields": {},
                        "no_empty_fields": [
                            "Application",
                            "Product Type",
                            "CAN Baud Rate",
                            "Connector",
                            "Connector 2",
                            "Software Version",
                            "Maximum Operating Temperature",
                            "Minimum Operating Temperature",
                        ],
                    },
                },
                "N": {
                    "Main Info": {
                        "equal_fields": {
                            "Article Type Value": "AF; Finished Part",
                            "Saleable Item Type Value": "IE; Exchange Product",
                            "Title ID Value": "Front Radar Assembly; 3753",
                            "Use Status": "Series",
                            "Type Number": type_dropdown.value,
                        },
                        "no_empty_fields": [],
                    },
                    "Product": {
                        "equal_fields": {
                            "Product Group": "29LA",
                            "Product Group Disp": "Forward Looking Radar",
                            "Device Code": ["4963", "50013", "50016"],
                            "Saleable only to Customer(s)": "",
                            "Sales Channel Limitation": "AM only",
                            "Price Book": "BXA (air products)",
                            "Basic Part Number": part,
                        },
                        "no_empty_fields": [],
                    },
                    "Special Characteristics": {
                        "equal_fields": {},
                        "no_empty_fields": ["Number of S/C", "Number of C/C"],
                    },
                    "Classification": {
                        "equal_fields": {},
                        "no_empty_fields": [
                            "Application",
                            "Product Type",
                            "CAN Baud Rate",
                            "Connector",
                            "Connector 2",
                            "Software Version",
                            "Maximum Operating Temperature",
                            "Minimum Operating Temperature",
                        ],
                    },
                },
                "SC": {
                    "Main Info": {
                        "equal_fields": {
                            "Article Type Value": "AF; Finished Part",
                            "Saleable Item Type Value": "IP; Product",
                            "Title ID Value": "Front Radar Assembly; 3753",
                            "Use Status": "Series",
                            "Type Number": type_dropdown.value,
                        },
                        "no_empty_fields": [],
                    },
                    "Product": {
                        "equal_fields": {
                            "Product Group": "29LA",
                            "Product Group Disp": "Forward Looking Radar",
                            "Device Code": ["4963", "50013", "50016"],
                            "Saleable only to Customer(s)": "All Customers",
                            "Sales Channel Limitation": "AM only",
                            "Price Book": "BXA (air products)",
                            "Basic Part Number": part,
                        },
                        "no_empty_fields": [],
                    },
                    "Special Characteristics": {
                        "equal_fields": {},
                        "no_empty_fields": ["Number of S/C", "Number of C/C"],
                    },
                    "Classification": {
                        "equal_fields": {},
                        "no_empty_fields": [
                            "Application",
                            "Product Type",
                            "CAN Baud Rate",
                            "Connector",
                            "Connector 2",
                            "Software Version",
                            "Maximum Operating Temperature",
                            "Minimum Operating Temperature",
                        ],
                    },
                },
            },
            "FLC-25": {
                "R": {
                    "Main Info": {
                        "equal_fields": {
                            "Article Type Value": "AF; Finished Part",
                            "Saleable Item Type Value": "IP; Product",
                            "Title ID Value": "Camera; 2687",
                            "Use Status": "Series",
                            "Type Number": type_dropdown.value,
                        },
                        "no_empty_fields": [],
                    },
                    "Product": {
                        "equal_fields": {
                            "Product Group": "29LL",
                            "Product Group Disp": "Forward Looking Camera",
                            "Device Code": "50024",
                            "Sales Channel Limitation": "OEM only",
                            "Price Book": "BXA (air products)",
                            "Basic Part Number": part,
                        },
                        "no_empty_fields": ["Saleable only to Customer(s)"],
                    },
                    "Special Characteristics": {
                        "equal_fields": {},
                        "no_empty_fields": ["Number of S/C", "Number of C/C"],
                    },
                    "Classification": {
                        "equal_fields": {},
                        "no_empty_fields": [
                            "Application",
                            "ECU Type",
                            "CAN Baud Rate",
                            "Camera Angle",
                            "Vehicle Connector",
                            "Video Connector",
                            "Camera Cable Length",
                            "Software Version",
                        ],
                    },
                },
                "X": {
                    "Main Info": {
                        "equal_fields": {
                            "Article Type Value": "AF; Finished Part",
                            "Saleable Item Type Value": "N; Not for external sale",
                            "Title ID Value": "Camera; 2687",
                            "Use Status": "Series",
                            "Type Number": type_dropdown.value,
                        },
                        "no_empty_fields": [],
                    },
                    "Product": {
                        "equal_fields": {
                            "Product Group": "29LL",
                            "Product Group Disp": "Forward Looking Camera",
                            "Device Code": "50024",
                            "Sales Channel Limitation": "Inter company only",
                            "Price Book": "BXA (air products)",
                            "Basic Part Number": part,
                        },
                        "no_empty_fields": [],
                    },
                    "Special Characteristics": {
                        "equal_fields": {},
                        "no_empty_fields": ["Number of S/C", "Number of C/C"],
                    },
                    "Classification": {
                        "equal_fields": {},
                        "no_empty_fields": [
                            "Application",
                            "ECU Type",
                            "CAN Baud Rate",
                            "Camera Angle",
                            "Vehicle Connector",
                            "Video Connector",
                            "Camera Cable Length",
                            "Software Version",
                        ],
                    },
                },
                "SC": {
                    "Main Info": {
                        "equal_fields": {
                            "Article Type Value": "AF; Finished Part",
                            "Saleable Item Type Value": "IP; Product",
                            "Title ID Value": "Camera; 2687",
                            "Use Status": "Series",
                            "Type Number": type_dropdown.value,
                        },
                        "no_empty_fields": [],
                    },
                    "Product": {
                        "equal_fields": {
                            "Product Group": "29LL",
                            "Product Group Disp": "Forward Looking Camera",
                            "Device Code": "50024",
                            "Sales Channel Limitation": "AM only",
                            "Price Book": "BXA (air products)",
                            "Basic Part Number": part,
                        },
                        "no_empty_fields": ["Saleable only to Customer(s)"],
                    },
                    "Special Characteristics": {
                        "equal_fields": {},
                        "no_empty_fields": ["Number of S/C", "Number of C/C"],
                    },
                    "Classification": {
                        "equal_fields": {},
                        "no_empty_fields": [
                            "Application",
                            "ECU Type",
                            "CAN Baud Rate",
                            "Camera Angle",
                            "Vehicle Connector",
                            "Video Connector",
                            "Camera Cable Length",
                            "Software Version",
                        ],
                    },
                },
            },
            "FLR-25": {
                "R": {
                    "Main Info": {
                        "equal_fields": {
                            "Article Type Value": "AF; Finished Part",
                            "Saleable Item Type Value": "IP; Product",
                            "Title ID Value": "Front Radar Assembly; 3753",
                            "Use Status": "Series",
                            "Type Number": type_dropdown.value,
                        },
                        "no_empty_fields": [],
                    },
                    "Product": {
                        "equal_fields": {
                            "Product Group": "29LA",
                            "Product Group Disp": "Forward Looking Radar",
                            "Device Code": ["50007", "50022"],
                            "Sales Channel Limitation": "OEM only",
                            "Price Book": "BXA (air products)",
                            "Basic Part Number": part,
                        },
                        "no_empty_fields": ["Saleable only to Customer(s)"],
                    },
                    "Special Characteristics": {
                        "equal_fields": {},
                        "no_empty_fields": ["Number of S/C", "Number of C/C"],
                    },
                    "Classification": {
                        "equal_fields": {},
                        "no_empty_fields": [
                            "Application",
                            "Product Type",
                            "CAN Baud Rate",
                            "Connector",
                            "Connector 2",
                            "Software Version",
                            "Maximum Operating Temperature",
                            "Minimum Operating Temperature",
                        ],
                    },
                },
                "X": {
                    "Main Info": {
                        "equal_fields": {
                            "Article Type Value": "AF; Finished Part",
                            "Saleable Item Type Value": "N; Not for external sale",
                            "Title ID Value": "Front Radar Assembly; 3753",
                            "Use Status": "Series",
                            "Type Number": type_dropdown.value,
                        },
                        "no_empty_fields": [],
                    },
                    "Product": {
                        "equal_fields": {
                            "Product Group": "29LA",
                            "Product Group Disp": "Forward Looking Radar",
                            "Device Code": ["50007", "50022"],
                            "Sales Channel Limitation": "Inter company only",
                            "Price Book": "BXA (air products)",
                            "Basic Part Number": part,
                        },
                        "no_empty_fields": [],
                    },
                    "Special Characteristics": {
                        "equal_fields": {},
                        "no_empty_fields": ["Number of S/C", "Number of C/C"],
                    },
                    "Classification": {
                        "equal_fields": {},
                        "no_empty_fields": [
                            "Application",
                            "Product Type",
                            "CAN Baud Rate",
                            "Connector",
                            "Connector 2",
                            "Software Version",
                            "Maximum Operating Temperature",
                            "Minimum Operating Temperature",
                        ],
                    },
                },
                "SC": {
                    "Main Info": {
                        "equal_fields": {
                            "Article Type Value": "AF; Finished Part",
                            "Saleable Item Type Value": "IP; Product",
                            "Title ID Value": "Front Radar Assembly; 3753",
                            "Use Status": "Series",
                            "Type Number": type_dropdown.value,
                        },
                        "no_empty_fields": [],
                    },
                    "Product": {
                        "equal_fields": {
                            "Product Group": "29LA",
                            "Product Group Disp": "Forward Looking Radar",
                            "Device Code": ["50007", "50022"],
                            "Sales Channel Limitation": "AM only",
                            "Price Book": "BXA (air products)",
                            "Basic Part Number": part,
                        },
                        "no_empty_fields": ["Saleable only to Customer(s)"],
                    },
                    "Special Characteristics": {
                        "equal_fields": {},
                        "no_empty_fields": ["Number of S/C", "Number of C/C"],
                    },
                    "Classification": {
                        "equal_fields": {},
                        "no_empty_fields": [
                            "Application",
                            "Product Type",
                            "CAN Baud Rate",
                            "Connector",
                            "Connector 2",
                            "Software Version",
                            "Maximum Operating Temperature",
                            "Minimum Operating Temperature",
                        ],
                    },
                },
            },
        }

        # Return the required attributes based on type_dropdown and letter_identifier
        return required_attributes.get(type_dropdown.value, {}).get(
            letter_identifier, {}
        )

    def get_attribute(self, driver, attr):
        data_locator_fields = [
            "Article Type Value",
            "Saleable Item Type Value",
            "Title ID Value",
            "Use Status",
            "Type Number",
            "Product Group",
            "Product Group Disp",
            "Product Group Frozen",
            "Device Code",
            "Sales Channel Limitation",
            "Price Book",
            "Basic Part Number",
            "Number of S/C",
            "Number of C/C",
        ]
        if attr in data_locator_fields:
            element_value = driver.find_element(
                By.CSS_SELECTOR, f"span[data-locator='{attr}']"
            ).get_attribute("value")
        else:
            li_elements = driver.find_elements(
                By.XPATH,
                f'//label[@class="sw-property sw-component sw-row sw-readOnly"]/span[@class="sw-property-name" and text()="{attr}"]/..//ul//li',
            )
            element_value = ""
            if li_elements:
                element_value = ", ".join(li.text for li in li_elements)
        return element_value if element_value is not None else ""

    def get_attribute_values(
        self, driver, part_number, type_dropdown, letter_identifier
    ):
        elements = {
            "Parameters": [],
            "Expected Value": [],
            "Actual Value": [],
            "Should be Equal": [],
        }

        try:
            attrs = self.attributes(part_number, type_dropdown, letter_identifier)
        except AttributeError:
            return pd.DataFrame(elements)

        for section in attrs:
            # Process attributes with expected values
            for attr, expected_value in attrs[section]["equal_fields"].items():
                try:
                    element = self.get_attribute(driver, attr)
                    elements["Parameters"].append(attr)
                    elements["Expected Value"].append(expected_value)
                    elements["Actual Value"].append(element)
                    elements["Should be Equal"].append(True)
                except NoSuchElementException:
                    elements["Parameters"].append(attr)
                    elements["Expected Value"].append(expected_value)
                    elements["Actual Value"].append("")
                    elements["Should be Equal"].append(True)

            for attr in attrs[section]["no_empty_fields"]:
                try:
                    element = self.get_attribute(driver, attr)
                    elements["Parameters"].append(attr)
                    elements["Expected Value"].append("")
                    elements["Actual Value"].append(element)
                    elements["Should be Equal"].append(False)
                except NoSuchElementException:
                    elements["Parameters"].append(attr)
                    elements["Expected Value"].append(expected_value)
                    elements["Actual Value"].append("")
                    elements["Should be Equal"].append(False)

        df = pd.DataFrame(elements)

        df["Expected Value"] = df.apply(
            lambda row: (
                format_device_code(row["Expected Value"])
                if row["Parameters"] == "Device Code"
                else row["Expected Value"]
            ),
            axis=1,
        )
        df["Expected Value"] = df.apply(modify_expected_value, axis=1)

        styled_df = (
            df.style.apply(highlight_row, axis=1)
            .hide(subset=["Should be Equal"], axis=1)
            .set_table_styles(
                [
                    {
                        "selector": "td:nth-child(2)",
                        "props": [("width", "250px")],
                    },  # 'Parameters' width
                    {
                        "selector": "td:nth-child(3)",
                        "props": [("width", "300px")],
                    },  # 'Expected Value' width
                    {
                        "selector": "td:nth-child(5)",
                        "props": [("width", "100px")],
                    },  # 'Should be Equal' width
                ]
            )
        )
        msg_pdf = f"\nResults for {part_number}:\n"
        print(f"\033[1m{msg_pdf}\033[0m")
        display(styled_df)
        return [{"msg": msg_pdf, "df": df}]

    def select_info(self, driver):
        """Select the Information tab."""
        wait = WebDriverWait(driver, 20)
        wait.until(
            EC.visibility_of_element_located(
                (By.XPATH, "//*[@id='main-view']//ul/li/a[text()='Information']")
            )
        )
        time.sleep(1 + 0.25 * self.extra_time_sec)
        driver.find_element(
            By.XPATH,
            "//li[a[@class='sw-row sw-aria-border sw-tab-title' and text()='Information']]",
        ).click()
        time.sleep(10 + self.extra_time_sec)

    def select_product_info(self, driver):
        wait = WebDriverWait(driver, 10)
        wait.until(
            EC.visibility_of_element_located(
                (
                    By.XPATH,
                    "//*[@id='main-view']//main/div/div[2]/details[1]/summary/div/div[text()='Product']",
                )
            )
        )
        time.sleep(1 + 0.25 * self.extra_time_sec)
        driver.find_element(
            By.XPATH, "//div[@class='sw-column sw-sectionTitle' and @title='Product']"
        ).click()
        time.sleep(10 + self.extra_time_sec)
        return self.get_sales_channel_limit(driver)

    def get_sales_channel_limit(self, driver):
        wait = WebDriverWait(driver, 10)
        sales_chann_limit = wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "span[data-locator='Sales Channel Limitation']")
            )
        ).get_attribute("value")
        return sales_chann_limit

    def select_content(self, driver):
        """Select the Content tab."""
        wait = WebDriverWait(driver, 20)
        wait.until(
            EC.visibility_of_element_located(
                (By.XPATH, "//*[@id='main-view']//ul/li/a[text()='Content']")
            )
        )
        time.sleep(1 + 0.25 * self.extra_time_sec)
        driver.find_element(
            By.XPATH,
            "//li[a[@class='sw-row sw-aria-border sw-tab-title' and text()='Content']]",
        ).click()
        time.sleep(20 + self.extra_time_sec)

    def select_products(self, driver):
        """Select the Products tab."""
        wait = WebDriverWait(driver, 20)
        wait.until(
            EC.visibility_of_element_located(
                (By.XPATH, "//*[@id='main-view']//ul/li/a[text()='Products']")
            )
        )
        time.sleep(1 + 0.25 * self.extra_time_sec)
        driver.find_element(
            By.XPATH,
            "//li[a[@class='sw-row sw-aria-border sw-tab-title' and text()='Products']]",
        ).click()
        time.sleep(12 + self.extra_time_sec)

    def validate_bracket_cover_number(self, part_number):
        """
        Validate if a part number contains a 'K' followed by numbers.

        Args:
            part_number (str): The part number to be validated.

        Returns:
            str or None: The letter identifier 'B' if the part number is valid, None otherwise.
        """
        # Regular expression pattern to search for a K followed by at least one number
        pattern = r"K\d+"
        # Search for matches in the part number
        matches = re.findall(pattern, part_number)

        # If there are matches, the validation is successful
        if len(matches) > 0:
            # self.letter_identifier = "B"
            return True  # for Bracket
        else:
            print(
                f"\033[48;5;208m\nValidation failed: Part number '{part_number}' does not contain 'K' followed by numbers.\n\033[0m"
            )
            # self.letter_identifier = None
            return False

    def validate_part_number(self, part_number):
        """
        Validate if a part number contains an 'X', 'R', 'N', or 'SC' followed by at least one number before and after the letter.

        Args:
            part_number (str): The part number to be validated.

        Returns:
            str or None: The letter identifier ('X', 'R', 'N', or 'SC') if the part number is valid, None otherwise.
        """
        # Regular expression pattern to search for an X, R, N or SC between by at least one number
        pattern = r"\d+(?:[XRN]|SC)\d+"
        # Search for matches in the part number
        matches = re.findall(pattern, part_number)

        letter_patterns = {"R": r"R", "X": r"X", "N": r"N", "SC": r"SC"}
        self.letter_identifier = next(
            (
                letter
                for letter, pattern in letter_patterns.items()
                if re.search(pattern, part_number)
            ),
            None,
        )
        # If there are matches, the validation is successful
        if len(matches) > 0:
            return self.letter_identifier
        else:
            print(
                f"\033[48;5;208m\nValidation failed: Part number '{part_number}' does not contain 'X', 'R', 'N' or 'SC' followed by at least one number before and after the letter.\n\033[0m"
            )
            return None

    def validate_partX_number(self, part_number):
        """
        Validate if a part number contains an 'X' followed by at least one number before and after the letter.
        """
        pattern = r"\d+X\d+"
        matches = re.findall(pattern, part_number)
        return part_number if matches else None

    def validate_type_number(self, part_number, type_number, type_dropdown):
        if type_dropdown.value == "FLR-25 Bracket":
            type_dropdown_value = "FLR-25"
        elif type_dropdown.value in ["FLC-25 Bracket", "FLC-25 Cover"]:
            type_dropdown_value = "FLC-25"
        else:
            type_dropdown_value = type_dropdown.value
        if type_number != type_dropdown_value:
            if type_number is None or type_number == "":
                print(
                    f"\033[1mThe Type Number attribute of the part {part_number} doesn't have a value\033[0m"
                )
            else:
                print(
                    f"\033[1;38;2;255;0;0m\nThe selected {type_dropdown_value} doesn't match the Type Number of the part {part_number}.\033[0m"
                )
                print(
                    f"\033[1mThe expected Type number attribute is {type_dropdown_value}, it was found: {type_number}\033[0m"
                )
            return False
        return True

    def get_identifier_bracket_cover(self, title_id_value, type_dropdown):
        if type_dropdown.value in ["FLR-25 Bracket", "FLC-25 Bracket", "FLC-25 Cover"]:
            if "bracket" in title_id_value.lower():
                self.letter_identifier = "B"
            elif "cover" in title_id_value.lower():
                self.letter_identifier = "Cover"
        return self.letter_identifier

    def get_headers_idx(self, driver, path_headers, headers):
        headers_idx = {}
        element_headers = driver.find_elements(By.XPATH, path_headers)

        for idx, element in enumerate(element_headers):
            header = element.find_element(By.XPATH, "./div[2]").get_attribute("title")
            if header in headers:
                headers_idx[header] = idx
                headers.remove(header)
            if not headers:
                break
        headers_idx = dict(sorted(headers_idx.items(), key=lambda x: x[1]))
        return headers_idx

    def create_title_tuples(self, values):
        """Transform the dictionary of values into a list of tuples with a specific order."""
        # Define the desired order of headers
        desired_order = ["Title ID Value", "Quantity", "Description", "ID"]
        # Initialize the list of tuples
        titles = []
        # Determine the number of rows (assuming all lists have the same length)
        num_rows = len(next(iter(values.values())))  # Take length from the first list
        # Iterate through each index of the rows
        for i in range(num_rows):
            # Construct the tuple in the desired order
            title_tuple = tuple(values[header][i] for header in desired_order)
            titles.append(title_tuple)
        return titles

    def get_titles(self, driver, div_elements, init_idx):
        """Get titles."""
        titles = []
        if not div_elements:
            return titles
        final_idx = init_idx + len(div_elements)
        headers_idx = self.get_headers_idx(
            driver,
            path_headers="//*[@id='occTreeTable']/div[2]/div[3]/div[1]/div",
            headers=["Title ID Value", "Quantity", "Description", "ID"],
        )
        time.sleep(3 + 0.5 * self.extra_time_sec)
        # Initialize the dictionary with empty lists
        values = {k: [] for k in headers_idx.keys()}

        actions = ActionChains(driver)
        previous_col_idx = 0

        def has_anchor_link(row_idx, col_idx):
            """Check if the cell contains an anchor link."""
            cell = self.wait_for_element(
                driver,
                (
                    By.XPATH,
                    f"//div[@aria-rowindex='{row_idx}']//div[@aria-colindex='{col_idx}']",
                ),
                timeout=10,
            )
            return cell.find_elements(By.XPATH, ".//a")

        def move_right(num_moves):
            """Move right by a number of steps with delay."""
            for j in range(num_moves):
                actions.send_keys(Keys.ARROW_RIGHT).perform()
                time.sleep(0.2 + 0.1 * self.extra_time_sec)
                if has_anchor_link(2, previous_col_idx + j + 2):
                    actions.send_keys(Keys.ARROW_RIGHT).perform()
                    time.sleep(0.4 + 0.1 * self.extra_time_sec)

        def click_and_check_title():
            """Click on the first element and check the title."""
            first_element = self.wait_for_element(
                driver,
                (By.XPATH, "//*[@id='occTreeTable_row2_col1']/div[1]/div"),
                timeout=10,
            )
            first_element.click()
            title_children = first_element.find_element(
                By.XPATH, ".//div[1]"
            ).get_attribute("title")
            if title_children == "Hide Children":
                actions.send_keys(Keys.ARROW_RIGHT).perform()
                time.sleep(0.2 + 0.1 * self.extra_time_sec)
            actions.send_keys(Keys.ARROW_RIGHT).perform()
            time.sleep(0.2 + 0.1 * self.extra_time_sec)
            return first_element

        def fetch_values(k, v):
            """Fetch values for the given header."""
            for i in range(init_idx, final_idx):
                val_element = self.wait_for_element(
                    driver,
                    (
                        By.XPATH,
                        f"//div[@aria-rowindex='{i+1}']//div[@aria-colindex='{v+2}']/div",
                    ),
                    timeout=10,
                )
                values[k].append(val_element.get_attribute("title").strip())

        for idx, (k, v) in enumerate(headers_idx.items()):
            if idx == 0:
                # Initial setup for the first column
                first_element = click_and_check_title()
                move_right(v)
            else:
                # Move to the right for the remaining columns
                num_moves = v - previous_col_idx
                move_right(num_moves)

            # Fetch values for the current header
            fetch_values(k, v)
            previous_col_idx = v

            if idx == len(headers_idx) - 1:
                # Prepare for the next action if needed
                click_and_check_title()

        titles = self.create_title_tuples(values)
        return titles

    # --- Document Validation Section ---
    def find_docs(self, driver, part_number):
        documents = {}
        last_aria_index = "0"
        while True:
            divs_docs = driver.find_elements(
                By.XPATH,
                ".//div[@class='aw-splm-tablePinnedContainer aw-splm-tablePinnedContainerLeft']//div[@class='aw-splm-tableScrollContents']/div",
            )

            if len(divs_docs) == 0:
                print(
                    f"\n\033[1;30;48;5;208mNot found elements in Documents section for {part_number}\033[0m"
                )
                break
            last_aria_rowindex = divs_docs[-1].get_attribute("aria-rowindex")
            if last_aria_rowindex == last_aria_index:
                break
            last_aria_index = last_aria_rowindex

            for div in divs_docs:
                aria_rowindex = div.get_attribute("aria-rowindex")
                title = driver.find_element(
                    By.XPATH,
                    f"//*[@id='ObjectSet_2_Provider_row{aria_rowindex}_col2']/div",
                ).get_attribute("title")
                if title not in documents:
                    headers_idx = self.get_headers_idx(
                        driver,
                        path_headers="//*[@id='ObjectSet_2_Provider']/div[2]/div[3]/div[1]/div",
                        headers=["Document Kind", "Lifecycle State"],
                    )
                    document_kind = driver.find_element(
                        By.XPATH,
                        f"//*[@id='ObjectSet_2_Provider_row{aria_rowindex}_col{headers_idx['Document Kind']+3}']/div",
                    ).get_attribute("title")
                    lifecycle_state = driver.find_element(
                        By.XPATH,
                        f"//*[@id='ObjectSet_2_Provider_row{aria_rowindex}_col{headers_idx['Lifecycle State']+3}']/div",
                    ).get_attribute("title")
                    documents[title] = {
                        "document_kind": document_kind,
                        "lifecycle_state": lifecycle_state,
                        "aria_rowindex": aria_rowindex,
                    }
            # Go to last div to allow load more elements:
            last_div = divs_docs[-1]
            last_div.find_element(By.XPATH, "./div[2]").click()
            time.sleep(2 + 0.5 * self.extra_time_sec)
        return documents

    def get_partX_name_in_content(self, driver):
        divs_docs = driver.find_elements(
            By.XPATH,
            ".//div[@class='aw-splm-tablePinnedContainer aw-splm-tablePinnedContainerLeft']//div[@class='aw-splm-tableScrollContents']/div",
        )
        elements_aria_level2 = [
            idx
            for idx, element in enumerate(divs_docs)
            if element.get_attribute("aria-level") == "2"
        ]
        headers_idx = self.get_headers_idx(
            driver,
            path_headers="//*[@id='occTreeTable']/div[2]/div[3]/div[1]/div",
            headers=["Title ID Value", "Quantity", "Description", "ID"],
        )
        for i in elements_aria_level2:
            title = driver.find_element(
                By.XPATH,
                f".//*[@id='occTreeTable_row{i+2}_col{headers_idx['ID']+2}']/div",
            ).get_attribute("title")
            has_partX = self.validate_partX_number(title.strip())
            if has_partX:
                return title.strip()
        print("It was not found a part X")
        return None

    # --- Aftermarket Section ---
    def documents_aftermarket_product(self, driver):
        """Select aftermarket product based on conditions."""
        time.sleep(1 + 0.25 * self.extra_time_sec)
        documents = {}
        self.click_element(
            driver,
            (By.XPATH, f"//*[@id='main-view']//main/div/details[5]/summary/div/div"),
        )
        time.sleep(1 + 0.25 * self.extra_time_sec)
        div_elements = driver.find_elements(
            By.XPATH, "//*[@id='ObjectSet_17_Provider']/div[2]/div[2]/div[2]/div/div"
        )
        idx = 0
        for div in div_elements:
            time.sleep(1 + 0.25 * self.extra_time_sec)
            idx += 1
            aria_rowindex = div.get_attribute("aria-rowindex")
            headers_idx = self.get_headers_idx(
                driver,
                path_headers="//*[@id='ObjectSet_17_Provider']/div[2]/div[3]/div[1]/div",
                headers=["Title ID Value", "Lifecycle State"],
            )
            title_id_value = self.wait_for_element(
                driver,
                (
                    By.XPATH,
                    f"//*[@id='ObjectSet_17_Provider_row{aria_rowindex}_col{headers_idx['Title ID Value']+3}']/div",
                ),
                timeout=10,
            ).get_attribute("title")
            document = self.wait_for_element(
                driver,
                (
                    By.XPATH,
                    f"//*[@id='ObjectSet_17_Provider_row{aria_rowindex}_col2']/div",
                ),
                timeout=10,
            ).get_attribute("title")
            lifecycle_state = self.wait_for_element(
                driver,
                (
                    By.XPATH,
                    f"//*[@id='ObjectSet_17_Provider_row{aria_rowindex}_col{headers_idx['Lifecycle State']+3}']/div",
                ),
                timeout=10,
            ).get_attribute("title")
            documents[idx] = {
                "Document": document,
                "Title_ID_value": title_id_value,
                "Lifecycle_State": lifecycle_state,
            }
        return documents

    def documents_aftermarket_product_for(self, driver):
        """Select aftermarket product based on conditions."""
        time.sleep(1 + 0.25 * self.extra_time_sec)
        documents = {}
        self.click_element(
            driver,
            (By.XPATH, f"//*[@id='main-view']//main/div/details[6]/summary/div/div"),
        )
        time.sleep(1 + 0.25 * self.extra_time_sec)
        div_elements = driver.find_elements(
            By.XPATH, "//*[@id='ObjectSet_18_Provider']/div[2]/div[2]/div[2]/div/div"
        )
        idx = 0
        for div in div_elements:
            time.sleep(1 + 0.25 * self.extra_time_sec)
            idx += 1
            aria_rowindex = div.get_attribute("aria-rowindex")
            headers_idx = self.get_headers_idx(
                driver,
                path_headers="//*[@id='ObjectSet_18_Provider']/div[2]/div[3]/div[1]/div",
                headers=["Title ID Value", "Lifecycle State"],
            )
            title_id_value = driver.find_element(
                By.XPATH,
                f"//*[@id='ObjectSet_18_Provider_row{aria_rowindex}_col{headers_idx['Title ID Value']+3}']/div",
            ).get_attribute("title")
            document = driver.find_element(
                By.XPATH,
                f"//*[@id='ObjectSet_18_Provider_row{aria_rowindex}_col2']/div",
            ).get_attribute("title")
            lifecycle_state = driver.find_element(
                By.XPATH,
                f"//*[@id='ObjectSet_18_Provider_row{aria_rowindex}_col{headers_idx['Lifecycle State']+3}']/div",
            ).get_attribute("title")
            documents[idx] = {
                "Document": document,
                "Title_ID_value": title_id_value,
                "Lifecycle_State": lifecycle_state,
            }
        return documents

    def process_missing_attributes(self, missing_elements):
        all_empty = True
        equal_fields_data = []
        empty_fields_data = []

        # Gather all equal_fields data
        for section_name, section_data in missing_elements.items():
            equal_fields = section_data["equal_fields"]
            if equal_fields:
                for field, values in equal_fields.items():
                    value = values["value"]
                    expected_value = values["expected_value"]
                    equal_fields_data.append(
                        {
                            "Section": section_name,
                            "Attribute": field,
                            "Value": value,
                            "Expected Value": expected_value,
                        }
                    )
                    all_empty = False

        # Gather all empty_fields data
        for section_name, section_data in missing_elements.items():
            no_empty_fields = section_data["no_empty_fields"]
            if no_empty_fields:
                for field in no_empty_fields:
                    empty_fields_data.append(
                        {"Section": section_name, "Attribute": field}
                    )
                    all_empty = False

        if all_empty:
            print(
                f"\033[1;30;48;5;34m\nSUCCESSFUL: \033[0m\033[1m\nAll attributes have the required data.\n\033[0m"
            )
            return

        # Print Equal Fields
        if equal_fields_data:
            equal_df = pd.DataFrame(equal_fields_data).replace({None: ""})
            print(
                f"\033[1m\033[48;5;208m\nThe following attributes doesn't match the expected value:\n\033[0m"
            )
            display(equal_df[["Attribute", "Value", "Expected Value"]])

        # Print Empty Fields
        if empty_fields_data:
            print(
                f"\033[1m\033[48;5;208m\nThe following attributes are empty, and they should have data:\n\033[0m"
            )
            for element in empty_fields_data:
                print("\033[1;38;2;255;0;0m- {}\033[0m".format(element["Attribute"]))
            print("\n")
        return

    # --- Schedule Section ---
    def search_schedule(self, driver, schedule_input):
        """Search for a schedule on the webpage."""
        search_element = self.wait_for_element(
            driver, (By.CSS_SELECTOR, "div.aw-search-globalSearchWidgetContainer")
        )
        search_element.click()
        time.sleep(1 + 0.25 * self.extra_time_sec)
        selector_element = self.wait_for_element(
            driver,
            (
                By.XPATH,
                "//*[@id='aw_navigation']//div[@class='aw-search-searchPreFilterPanel2']",
            ),
        )
        selector_element.click()
        self.wait_for_element(
            driver, (By.XPATH, "//div[@class='sw-cell-valName'][@title='Documents']")
        )
        time.sleep(1 + 0.25 * self.extra_time_sec)
        parts_element = self.wait_for_element(
            driver, (By.XPATH, "//div[@class='sw-cell-valName'][@title='Documents']")
        )
        parts_element.click()
        input_element = self.wait_for_element(
            driver,
            (
                By.XPATH,
                "//*[@id='aw_navigation']/div[2]/form/div[2]/div/div/div[1]/div/div[1]/div[2]/div/div/input",
            ),
        )
        input_element.send_keys(schedule_input)
        time.sleep(1 + 0.25 * self.extra_time_sec)
        search_button = self.wait_for_element(
            driver,
            (
                By.XPATH,
                "//*[@id='aw_navigation']/div[2]/form/div[2]/div/div/div[1]/div/div[1]/div[2]/div/div/div[2]",
            ),
        )
        search_button.click()

    def select_schedule_from_search(self, driver, schedule_value):
        """Select the document from the search results."""
        time.sleep(15 + self.extra_time_sec)
        li_elements = driver.find_elements(
            By.CSS_SELECTOR, "ul.aw-widgets-cellListWidget > li"
        )
        for i, li_element in enumerate(li_elements):
            time.sleep(1 + 0.25 * self.extra_time_sec)
            span_title_element = li_element.find_element(
                By.CSS_SELECTOR, "span.aw-widgets-cellListCellTitle"
            )
            title_text = span_title_element.get_attribute("title")
            title_clean = ",".join([part.strip() for part in title_text.split(",")])
            schedule_value_clean = ",".join(
                [part.strip() for part in schedule_value.split(",")]
            )
            if title_clean.upper() == schedule_value_clean.upper():
                driver.find_element(
                    By.XPATH, f"//*[@id='main-view']//ul/li[{i+1}]/div"
                ).click()
                time.sleep(1 + 0.25 * self.extra_time_sec)
                driver.find_element(
                    By.XPATH, f"//*[@id='main-view']//ul/li[{i+1}]/div/div[2]/div/div"
                ).click()
                return True
        print("\033[48;5;208mNo matching result for the Schedule Document\n\033[0m")
        return False
