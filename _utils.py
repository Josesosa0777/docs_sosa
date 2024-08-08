import re
import time
import warnings
from enum import Enum
from typing import Optional, List, Dict, Tuple, Union

import pandas as pd
from IPython.display import display
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class DriverType(Enum):
    """Enum for supported WebDriver types."""
    EDGE = "edge"
    FIREFOX = "firefox"
    CHROME = "chrome"


def format_device_code(expected_value: Union[str, List[str]]) -> str:
    """
    Format the device code for display.

    Args:
        expected_value (Union[str, List[str]]): The expected device code value.

    Returns:
        str: Formatted device code.
    """
    if isinstance(expected_value, list):
        return " or ".join(expected_value) if len(expected_value) > 1 else expected_value[0]
    return expected_value


def modify_expected_value(row: pd.Series) -> str:
    """
    Modify the expected value based on parameters.

    Args:
        row (pd.Series): A row of the DataFrame.

    Returns:
        str: Modified expected value.
    """
    parameter = row["Parameters"]
    if parameter in ["Number of S/C", "Number of C/C"]:
        return "(Digit, Not Blank)"
    elif parameter == "Application":
        return "(Type of Application, Not Blank)"
    else:
        return "(Not Blank)" if not row["Should be Equal"] else row["Expected Value"]


def highlight_row(x: pd.Series) -> List[str]:
    """
    Apply styles to the DataFrame rows.

    Args:
        x (pd.Series): A row of the DataFrame.

    Returns:
        List[str]: List of styles applied to the row.
    """
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
    """Utility class for common Selenium and data processing operations."""

    def __init__(self, extra_time: int = 0):
        """
        Initialize the Utils class.

        Args:
            extra_time (int): Additional time to wait for elements.
        """
        self.extra_time_sec = extra_time
        self.letter_identifier: Optional[str] = None

    def wait_for_element(self, driver: webdriver, locator: Tuple[By, str], timeout: int = 10):
        """Wait for element to be present on the page."""
        return WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located(locator)
        )

    def click_element(self, driver: webdriver, locator: Tuple[By, str]) -> None:
        """Click on the element identified by the locator."""
        element = self.wait_for_element(driver, locator)
        element.click()

    def open_url(self, url: str, driver_type: DriverType = DriverType.EDGE) -> webdriver:
        """Open the given URL in a new browser window."""
        with warnings.catch_warnings():
            warnings.filterwarnings(
                "ignore", message="selenium.webdriver.common.selenium_manager"
            )
            if driver_type == DriverType.EDGE:
                driver = webdriver.Edge()
            elif driver_type == DriverType.FIREFOX:
                driver = webdriver.Firefox()
            elif driver_type == DriverType.CHROME:
                driver = webdriver.Chrome()
            else:
                raise ValueError(
                    "Invalid driver type. Supported types are 'chrome', 'firefox', and 'edge'."
                )
        driver.get(url)
        time.sleep(10 + self.extra_time_sec)
        self.wait_for_element(driver, (By.TAG_NAME, "body"), timeout=50)
        return driver

    def close_session(self, driver: Optional[webdriver]) -> None:
        """Close the browser session."""
        if driver:
            try:
                driver.find_element(
                    By.XPATH,
                    f"//*[@id='main-view']/div/div[2]/div[2]/nav/div[1]/div[2]/button",
                ).click()
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
                print("Exception: Unable to close the browser.", e)
            time.sleep(4)
            driver.quit()

    def search_for_part(self, driver: webdriver, part_number: str) -> None:
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

    def select_from_search(self, driver: webdriver, part_number: str) -> bool:
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
                    By.XPATH, f"//*[@id='main-view']//ul/li[{i + 1}]/div/div[2]/div/div"
                ).click()
                return True
        print("\033[48;5;208mNo matching result for the Part Number\n\033[0m")
        return False

    def get_attr_from_part_search(self, driver: webdriver) -> Tuple[str, str, str, str, str]:
        """Get attributes from the part search results."""
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

    def expand_sections(self, driver: webdriver) -> None:
        """Expand sections on the webpage."""
        section_titles = ["Product", "Special Characteristics"]
        for section in section_titles:
            self.click_element(driver, (By.CSS_SELECTOR, f"div[title='{section}']"))

    def attributes(self, part_number: str, type_dropdown: str, letter_identifier: str) -> dict:
        """Get attributes required for validation based on part number, type and letter identifier."""
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
                            "Type Number": type_dropdown,
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
                # Continue for other identifiers "X", "N", "SC"...
            },
            # Continue for other types "FLR-21", "FLC-25", "FLR-25"...
        }

        return required_attributes.get(type_dropdown, {}).get(letter_identifier, {})

    def get_attribute(self, driver: webdriver, attr: str) -> str:
        """Get the value of a specific attribute from the webpage."""
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
            element_value = ", ".join(li.text for li in li_elements) if li_elements else ""
        return element_value

    def get_attribute_values(
        self, driver: webdriver, part_number: str, type_dropdown: str, letter_identifier: str
    ) -> List[Dict[str, Union[str, pd.DataFrame]]]:
        """Get and display the attribute values from the webpage."""
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
                    elements["Expected Value"].append("")
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

    def select_info(self, driver: webdriver) -> None:
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

    def select_product_info(self, driver: webdriver) -> str:
        """Select the Product tab and retrieve Sales Channel Limitation."""
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

    def get_sales_channel_limit(self, driver: webdriver) -> str:
        """Retrieve Sales Channel Limitation value."""
        wait = WebDriverWait(driver, 10)
        sales_chann_limit = wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "span[data-locator='Sales Channel Limitation']")
            )
        ).get_attribute("value")
        return sales_chann_limit

    def select_content(self, driver: webdriver) -> None:
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

    def select_products(self, driver: webdriver) -> None:
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

    def validate_bracket_cover_number(self, part_number: str) -> bool:
        """
        Validate if a part number contains a 'K' followed by numbers.

        Args:
            part_number (str): The part number to be validated.

        Returns:
            bool: True if the part number is valid, False otherwise.
        """
        pattern = r"K\d+"
        matches = re.findall(pattern, part_number)
        if matches:
            return True  # Valid bracket
        else:
            print(
                f"\033[48;5;208m\nValidation failed: Part number '{part_number}' does not contain 'K' followed by numbers.\n\033[0m"
            )
            return False

    def validate_part_number(self, part_number: str) -> Optional[str]:
        """
        Validate if a part number contains 'X', 'R', 'N', or 'SC' followed by numbers.

        Args:
            part_number (str): The part number to be validated.

        Returns:
            Optional[str]: The letter identifier if valid, None otherwise.
        """
        pattern = r"\d+(?:[XRN]|SC)\d+"
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
        if matches:
            return self.letter_identifier
        else:
            print(
                f"\033[48;5;208m\nValidation failed: Part number '{part_number}' does not contain 'X', 'R', 'N' or 'SC' followed by numbers.\n\033[0m"
            )
            return None

    def validate_partX_number(self, part_number: str) -> Optional[str]:
        """
        Validate if a part number contains an 'X' followed by numbers.

        Args:
            part_number (str): The part number to be validated.

        Returns:
            Optional[str]: The part number if valid, None otherwise.
        """
        pattern = r"\d+X\d+"
        matches = re.findall(pattern, part_number)
        return part_number if matches else None

    def validate_type_number(self, part_number: str, type_number: str, type_dropdown: str) -> bool:
        """Validate if the part number's type matches the selected type."""
        if type_dropdown == "FLR-25 Bracket":
            type_dropdown_value = "FLR-25"
        elif type_dropdown in ["FLC-25 Bracket", "FLC-25 Cover"]:
            type_dropdown_value = "FLC-25"
        else:
            type_dropdown_value = type_dropdown
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

    def get_identifier_bracket_cover(self, title_id_value: str, type_dropdown: str) -> Optional[str]:
        """Get the identifier for bracket or cover based on the title and type."""
        if type_dropdown in ["FLR-25 Bracket", "FLC-25 Bracket", "FLC-25 Cover"]:
            if "bracket" in title_id_value.lower():
                self.letter_identifier = "B"
            elif "cover" in title_id_value.lower():
                self.letter_identifier = "Cover"
        return self.letter_identifier

    def get_headers_idx(self, driver: webdriver, path_headers: str, headers: List[str]) -> Dict[str, int]:
        """Get the indices of headers in the table."""
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

    def create_title_tuples(self, values: Dict[str, List[str]]) -> List[Tuple[str, str, str, str]]:
        """Transform the dictionary of values into a list of tuples with a specific order."""
        desired_order = ["Title ID Value", "Quantity", "Description", "ID"]
        titles = []
        num_rows = len(next(iter(values.values())))
        for i in range(num_rows):
            title_tuple = tuple(values[header][i] for header in desired_order)
            titles.append(title_tuple)
        return titles

    def get_titles(self, driver: webdriver, div_elements: List[webdriver], init_idx: int) -> List[Tuple[str, str, str, str]]:
        """Get titles from the webpage."""
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
        values = {k: [] for k in headers_idx.keys()}

        actions = ActionChains(driver)
        previous_col_idx = 0

        def has_anchor_link(row_idx: int, col_idx: int) -> bool:
            """Check if the cell contains an anchor link."""
            cell = self.wait_for_element(
                driver,
                (
                    By.XPATH,
                    f"//div[@aria-rowindex='{row_idx}']//div[@aria-colindex='{col_idx}']",
                ),
                timeout=10,
            )
            return bool(cell.find_elements(By.XPATH, ".//a"))

        def move_right(num_moves: int) -> None:
            """Move right by a number of steps with delay."""
            for j in range(num_moves):
                actions.send_keys(Keys.ARROW_RIGHT).perform()
                time.sleep(0.2 + 0.1 * self.extra_time_sec)
                if has_anchor_link(2, previous_col_idx + j + 2):
                    actions.send_keys(Keys.ARROW_RIGHT).perform()
                    time.sleep(0.4 + 0.1 * self.extra_time_sec)

        def click_and_check_title() -> webdriver:
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

        def fetch_values(k: str, v: int) -> None:
            """Fetch values for the given header."""
            for i in range(init_idx, final_idx):
                val_element = self.wait_for_element(
                    driver,
                    (
                        By.XPATH,
                        f"//div[@aria-rowindex='{i + 1}']//div[@aria-colindex='{v + 2}']/div",
                    ),
                    timeout=10,
                )
                values[k].append(val_element.get_attribute("title").strip())

        for idx, (k, v) in enumerate(headers_idx.items()):
            if idx == 0:
                first_element = click_and_check_title()
                move_right(v)
            else:
                num_moves = v - previous_col_idx
                move_right(num_moves)
            fetch_values(k, v)
            previous_col_idx = v
            if idx == len(headers_idx) - 1:
                click_and_check_title()

        titles = self.create_title_tuples(values)
        return titles

    def find_docs(self, driver: webdriver, part_number: str) -> Dict[str, Dict[str, Union[str, int]]]:
        """Find documents associated with the part."""
        documents = {}
        last_aria_index = "0"
        while True:
            divs_docs = driver.find_elements(
                By.XPATH,
                ".//div[@class='aw-splm-tablePinnedContainer aw-splm-tablePinnedContainerLeft']//div[@class='aw-splm-tableScrollContents']/div",
            )
            if not divs_docs:
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
                        f"//*[@id='ObjectSet_2_Provider_row{aria_rowindex}_col{headers_idx['Document Kind'] + 3}']/div",
                    ).get_attribute("title")
                    lifecycle_state = driver.find_element(
                        By.XPATH,
                        f"//*[@id='ObjectSet_2_Provider_row{aria_rowindex}_col{headers_idx['Lifecycle State'] + 3}']/div",
                    ).get_attribute("title")
                    documents[title] = {
                        "document_kind": document_kind,
                        "lifecycle_state": lifecycle_state,
                        "aria_rowindex": aria_rowindex,
                    }
            last_div = divs_docs[-1]
            last_div.find_element(By.XPATH, "./div[2]").click()
            time.sleep(2 + 0.5 * self.extra_time_sec)
        return documents

    def get_partX_name_in_content(self, driver: webdriver) -> Optional[str]:
        """Get the part X name from the content."""
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
                f".//*[@id='occTreeTable_row{i + 2}_col{headers_idx['ID'] + 2}']/div",
            ).get_attribute("title")
            has_partX = self.validate_partX_number(title.strip())
            if has_partX:
                return title.strip()
        print("It was not found a part X")
        return None

    def documents_aftermarket_product(self, driver: webdriver) -> Dict[int, Dict[str, str]]:
        """Retrieve documents for aftermarket product."""
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
        for idx, div in enumerate(div_elements, start=1):
            time.sleep(1 + 0.25 * self.extra_time_sec)
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
                    f"//*[@id='ObjectSet_17_Provider_row{aria_rowindex}_col{headers_idx['Title ID Value'] + 3}']/div",
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
                    f"//*[@id='ObjectSet_17_Provider_row{aria_rowindex}_col{headers_idx['Lifecycle State'] + 3}']/div",
                ),
                timeout=10,
            ).get_attribute("title")
            documents[idx] = {
                "Document": document,
                "Title_ID_value": title_id_value,
                "Lifecycle_State": lifecycle_state,
            }
        return documents

    def documents_aftermarket_product_for(self, driver: webdriver) -> Dict[int, Dict[str, str]]:
        """Retrieve documents for aftermarket product."""
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
        for idx, div in enumerate(div_elements, start=1):
            time.sleep(1 + 0.25 * self.extra_time_sec)
            aria_rowindex = div.get_attribute("aria-rowindex")
            headers_idx = self.get_headers_idx(
                driver,
                path_headers="//*[@id='ObjectSet_18_Provider']/div[2]/div[3]/div[1]/div",
                headers=["Title ID Value", "Lifecycle State"],
            )
            title_id_value = driver.find_element(
                By.XPATH,
                f"//*[@id='ObjectSet_18_Provider_row{aria_rowindex}_col{headers_idx['Title ID Value'] + 3}']/div",
            ).get_attribute("title")
            document = driver.find_element(
                By.XPATH,
                f"//*[@id='ObjectSet_18_Provider_row{aria_rowindex}_col2']/div",
            ).get_attribute("title")
            lifecycle_state = driver.find_element(
                By.XPATH,
                f"//*[@id='ObjectSet_18_Provider_row{aria_rowindex}_col{headers_idx['Lifecycle State'] + 3}']/div",
            ).get_attribute("title")
            documents[idx] = {
                "Document": document,
                "Title_ID_value": title_id_value,
                "Lifecycle_State": lifecycle_state,
            }
        return documents

    def process_missing_attributes(self, missing_elements: Dict[str, Dict[str, Dict[str, Union[str, None]]]]) -> None:
        """Process and display missing attributes."""
        all_empty = True
        equal_fields_data = []
        empty_fields_data = []

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

        if equal_fields_data:
            equal_df = pd.DataFrame(equal_fields_data).replace({None: ""})
            print(
                f"\033[1m\033[48;5;208m\nThe following attributes doesn't match the expected value:\n\033[0m"
            )
            display(equal_df[["Attribute", "Value", "Expected Value"]])

        if empty_fields_data:
            print(
                f"\033[1m\033[48;5;208m\nThe following attributes are empty, and they should have data:\n\033[0m"
            )
            for element in empty_fields_data:
                print("\033[1;38;2;255;0;0m- {}\033[0m".format(element["Attribute"]))
            print("\n")

    def search_schedule(self, driver: webdriver, schedule_input: str) -> None:
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

    def select_schedule_from_search(self, driver: webdriver, schedule_value: str) -> bool:
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
                    By.XPATH, f"//*[@id='main-view']//ul/li[{i + 1}]/div"
                ).click()
                time.sleep(1 + 0.25 * self.extra_time_sec)
                driver.find_element(
                    By.XPATH, f"//*[@id='main-view']//ul/li[{i + 1}]/div/div[2]/div/div"
                ).click()
                return True
        print("\033[48;5;208mNo matching result for the Schedule Document\n\033[0m")
        return False
