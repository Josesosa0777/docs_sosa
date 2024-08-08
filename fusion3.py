import re
import time
from typing import List, Optional, Tuple, Union

import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.remote.webdriver import WebDriver
from IPython.display import display, Markdown

from enums import ElementType, DropdownType, LetterIdentifier


class Fusion3:
    """Class containing useful functions and methods for automation of Fusion 3.0."""

    def __init__(self, utils, extra_time: Optional[Union[int, float]] = 0):
        """
        Initialize the Fusion3 class.

        Args:
            utils: An instance of a utility class for interacting with the web page.
            extra_time: Optional additional time (in seconds) to adjust delays for actions.
        """
        self.extra_time_sec = int(extra_time)
        self.utils = utils
        self.is_am_bracket = False

    def initialize_main_params(
        self,
        part_number: str,
        title_id_value: str,
        driver: WebDriver,
        child_div_path: str,
        letter_identifier: LetterIdentifier,
        chbx_bracket: Optional[bool],
        chbx_resistor: Optional[bool],
        type_dropdown,
    ) -> None:
        """
        Initialize main parameters and validate conditions based on the provided values.

        Args:
            part_number: The part number to initialize.
            title_id_value: The title ID value to check for specific keywords.
            driver: The Selenium WebDriver instance used for automation.
            child_div_path: XPath expression to locate child div elements.
            letter_identifier: Identifier for the type of product, e.g., "R" or "SC".
            chbx_bracket: Boolean indicating whether the bracket checkbox is checked.
            chbx_resistor: Boolean indicating whether the resistor checkbox is checked.
            type_dropdown: The dropdown selection representing the type of element.

        Raises:
            ValueError: If the title ID value does not match the expected conditions for the selected type.
        """
        self.part_number = part_number
        self.title_id_value = title_id_value
        self.driver = driver
        self.child_div_path = child_div_path
        self.letter_identifier = letter_identifier
        self.chbx_bracket = chbx_bracket
        self.chbx_resistor = chbx_resistor
        self.type_dropdown = type_dropdown
        self.radar_in_title = "radar" in self.title_id_value.lower()
        self.camera_in_title = "camera" in self.title_id_value.lower()
        self.bracket_in_title = "bracket assembly" in self.title_id_value.lower()
        self.cover_in_title = "cover" in self.title_id_value.lower()
        if type_dropdown.value in ["FLR-25 Bracket", "FLC-25 Bracket"]:
            if not self.bracket_in_title:
                raise ValueError(
                    "No specific required element defined for Bracket Assembly. Expected to have 'Bracket Assembly' in the Title ID Value attribute"
                )
            self.element_name = "Bracket"
        elif type_dropdown.value in ["FLC-25 Cover"]:
            if not self.cover_in_title:
                raise ValueError(
                    "No specific required element defined for Cover. Expected to have 'Cover' in the Title ID Value attribute"
                )
            self.element_name = "Cover"
        else:
            if not (self.radar_in_title or self.camera_in_title):
                raise ValueError(
                    "No specific required element defined for Camera or Radar."
                )
            self.element_name = "Radar" if self.radar_in_title else "Camera"

    def is_AM_bracket(self, driver):
        sales_chann_limit = self.utils.select_product_info(driver)
        if sales_chann_limit.lower().strip() == "am only":
            self.is_am_bracket = True
        else:
            self.is_am_bracket = False
        return self.is_am_bracket

    def expand_elements_get_titles(self):
        """Expand elements based on conditions."""
        time.sleep(0.5 + 0.25 * self.extra_time_sec)
        separator = self.driver.find_element(
            By.XPATH,
            "//*[@id='main-view']/div/div[2]/div[1]/div/div/div/div/div[3]/div/div/div[1]/div/div/main/div[2]",
        )
        # ActionChains instance
        actions = ActionChains(self.driver)
        element_location = self.driver.execute_script(
            "return arguments[0].getBoundingClientRect();", separator
        )
        viewport_width = self.driver.execute_script("return window.innerWidth;")
        element_x = element_location["left"]
        element_width = element_location["width"]
        move_x = min(250, viewport_width - element_x - element_width)
        actions.click_and_hold(separator).move_by_offset(move_x, 0).release().perform()
        div_elements = self.driver.find_elements(
            By.XPATH, f"{self.child_div_path}[@aria-level='2']"
        )
        titles = self.utils.get_titles(self.driver, div_elements, 2)

        if self.radar_in_title:
            if self.letter_identifier == "R":
                if self.chbx_bracket.value:
                    idx = next(
                        (
                            index
                            for index, (title, _, _, _) in enumerate(titles)
                            if "bracket" in title.lower()
                        ),
                        None,
                    )
                    if idx is not None:
                        self.utils.click_element(
                            self.driver,
                            (
                                By.XPATH,
                                f"//*[@id='occTreeTable_row{idx+3}_col1']/div/div/div[1]",
                            ),
                        )
                        time.sleep(4 + 0.5 * self.extra_time_sec)
                        bracket_elements = self.driver.find_elements(
                            By.XPATH, f"{self.child_div_path}[@aria-level='3']"
                        )
                        if bracket_elements:
                            titles.pop(idx)
                        time.sleep(3 + 0.5 * self.extra_time_sec)
                        child_titles = self.utils.get_titles(
                            self.driver, bracket_elements, idx + 3
                        )
                        titles.extend(child_titles)
            if self.letter_identifier == "SC":
                idx = next(
                    (
                        index
                        for index, (title, _, _, _) in enumerate(titles)
                        if "radar" in title.lower()
                    ),
                    None,
                )
                if idx is not None:
                    partX = self.utils.click_element(
                        self.driver,
                        (
                            By.XPATH,
                            f"//*[@id='occTreeTable_row{idx+3}_col1']/div/div/div[1]",
                        ),
                    )
                    time.sleep(4 + 0.5 * self.extra_time_sec)
                    partX_elements = self.driver.find_elements(
                        By.XPATH, f"{self.child_div_path}[@aria-level='3']"
                    )
                    if partX_elements:
                        titles.pop(idx)
                    time.sleep(3 + 0.5 * self.extra_time_sec)
                    child_titles = self.utils.get_titles(
                        self.driver, partX_elements, idx + 3
                    )
                    titles.extend(child_titles)
        if self.camera_in_title:
            if self.letter_identifier == "R":
                if self.chbx_bracket.value:
                    idx = next(
                        (
                            index
                            for index, (title, _, _, _) in enumerate(titles)
                            if "bracket" in title.lower()
                        ),
                        None,
                    )
                    if idx is not None:
                        self.utils.click_element(
                            self.driver,
                            (
                                By.XPATH,
                                f"//*[@id='occTreeTable_row{idx+3}_col1']/div/div/div[1]",
                            ),
                        )
                        time.sleep(4 + 0.5 * self.extra_time_sec)
                        bracket_elements = self.driver.find_elements(
                            By.XPATH, f"{self.child_div_path}[@aria-level='3']"
                        )
                        if bracket_elements:
                            titles.pop(idx)
                        time.sleep(3 + 0.5 * self.extra_time_sec)
                        child_titles = self.utils.get_titles(
                            self.driver, bracket_elements, idx + 3
                        )
                        titles.extend(child_titles)
            if self.letter_identifier == "SC":
                idx = next(
                    (
                        index
                        for index, (title, _, _, _) in enumerate(titles)
                        if "camera" in title.lower()
                    ),
                    None,
                )
                if idx is not None:
                    self.utils.click_element(
                        self.driver,
                        (
                            By.XPATH,
                            f"//*[@id='occTreeTable_row{idx+3}_col1']/div/div/div[1]",
                        ),
                    )
                    time.sleep(4 + 0.5 * self.extra_time_sec)
                    partX_elements = self.driver.find_elements(
                        By.XPATH, f"{self.child_div_path}[@aria-level='3']"
                    )
                    if partX_elements:
                        titles.pop(idx)
                    time.sleep(3 + 0.5 * self.extra_time_sec)
                    child_titles = self.utils.get_titles(
                        self.driver, partX_elements, idx + 3
                    )
                    titles.extend(child_titles)
        if self.bracket_in_title:
            if self.type_dropdown.value == "FLR-25 Bracket":
                idx = next(
                    (
                        index
                        for index, (title, _, _, _) in enumerate(titles)
                        if "bracket" in title.lower()
                    ),
                    None,
                )
                if idx is not None:
                    self.utils.click_element(
                        self.driver,
                        (
                            By.XPATH,
                            f"//*[@id='occTreeTable_row{idx+3}_col1']/div/div/div[1]",
                        ),
                    )
                    time.sleep(4 + 0.5 * self.extra_time_sec)
                    bracket_elements = self.driver.find_elements(
                        By.XPATH, f"{self.child_div_path}[@aria-level='3']"
                    )
                    if bracket_elements:
                        titles.pop(idx)
                    time.sleep(3 + 0.5 * self.extra_time_sec)
                    child_titles = self.utils.get_titles(
                        self.driver, bracket_elements, idx + 3
                    )
                    titles.extend(child_titles)
        actions.click_and_hold(separator).move_by_offset(-move_x, 0).release().perform()
        return titles

    # --- BOM Structure Section ---
    def get_required_elements(self, element_name):
        """Get required elements based on element name and letter identifier."""
        elements = []
        if element_name == "Radar":
            if self.letter_identifier == "X":
                elements = ["Dataset", "Software", "Boot Software"] + (
                    ["with CAN termination"]
                    if self.chbx_resistor.value
                    else ["wo CAN termination"]
                )
            elif self.letter_identifier == "R":
                elements = (
                    ["Dataset", "Software", "Boot Software"]
                    + (["Screw", "Bracket", "Nut"] if self.chbx_bracket.value else [])
                    + (
                        ["with CAN termination"]
                        if self.chbx_resistor.value
                        else ["wo CAN termination"]
                    )
                )
            elif self.letter_identifier == "SC":
                elements = ["Dataset", "Software", "Boot Software"] + (
                    ["with CAN termination"]
                    if self.chbx_resistor.value
                    else ["wo CAN termination"]
                )
        if element_name == "Camera":
            if self.letter_identifier in ("X", "SC"):
                elements = ["Dataset", "Software", "Boot Software", "Camera"]
            elif self.letter_identifier == "R":
                elements = ["Dataset", "Software", "Boot Software", "Camera"]
        if element_name == "Bracket" and self.type_dropdown.value == "FLR-25 Bracket":
            elements = (
                ["Screw", "Bracket", "Nut"]
                if self.is_am_bracket
                else ["Bracket", "Nut"]
            )
        if element_name == "Bracket" and self.type_dropdown.value == "FLC-25 Bracket":
            elements = ["Bracket", "Bracket Assembly", "Worm", "Tape"]
        if element_name == "Cover" and self.type_dropdown.value == "FLC-25 Cover":
            elements = ["Cover"]
        return elements

    def get_bom_in_content(self):
        """Get titles in Content tab."""
        titles = self.expand_elements_get_titles()
        time.sleep(1 + 0.25 * self.extra_time_sec)

        required_elements = self.get_required_elements(self.element_name)

        missing_elements = [
            element
            for element in required_elements
            if not any(element in title for title, _, _, _ in titles)
        ]

        # Create a dictionary to store the data
        data = {
            "ID": [],
            "Title": [],
            "Description": [],
            "Quantity": [],
            "Expected_value": [],
        }

        if self.radar_in_title:
            if self.letter_identifier in ("R", "SC", "X"):
                # Validate the presence of 'Dataset' in missing_elements
                if "Dataset" in missing_elements:
                    # Count how many times 'Software' appears in titles
                    count_software = sum(
                        1
                        for title, _, description, _ in titles
                        if "software" in title.lower()
                    )
                    # Check if there are at least 2 'Software' elements and at least one containing 'DATASET' in the description
                    if count_software >= 2 and any(
                        (
                            "DATASET" in description.upper()
                            or "DATA SET" in description.upper()
                        )
                        for title, _, description, _ in titles
                        if "software" in title.lower()
                    ):
                        missing_elements.remove("Dataset")

                # Validate the presence of 'Software' in missing_elements
                boot_software_found = any(
                    "software" in title.lower()
                    and "boot software" in description.lower()
                    for title, _, description, _ in titles
                )
                if not boot_software_found:
                    missing_elements.append("Boot Software")
                else:
                    if "Boot Software" in missing_elements:
                        missing_elements.remove("Boot Software")
                # Check if there is an element containing "Software" in the title and matches the defined pattern, and does not have "Boot" in the description
                pattern = r"VER FU\d{6}"
                software_ver_found = any(
                    "software" in title.lower() and re.search(pattern, description)
                    for title, _, description, _ in titles
                    if "boot software" not in description.lower()
                )
                if not software_ver_found and not "Software" in missing_elements:
                    missing_elements.append("Software")
                if self.chbx_resistor.value:
                    can_termination = any(
                        "radar" in title.lower()
                        and "with can termination" in description.lower()
                        and id_.upper() == "K218450H002"
                        for title, _, description, id_ in titles
                    )
                    if not can_termination:
                        if "with CAN termination" in missing_elements:
                            missing_elements.remove("with CAN termination")
                        missing_elements.append(
                            "Radar with CAN termination (K218450H002)"
                        )
                    else:
                        if "with CAN termination" in missing_elements:
                            missing_elements.remove("with CAN termination")
                else:
                    can_termination = any(
                        "radar" in title.lower()
                        and "wo can termination" in description.lower()
                        and id_.upper() == "K188333H002"
                        for title, _, description, id_ in titles
                    )
                    if not can_termination:
                        if "wo CAN termination" in missing_elements:
                            missing_elements.remove("wo CAN termination")
                        missing_elements.append(
                            "Radar without CAN termination (K188333H002)"
                        )
                    else:
                        if "wo CAN termination" in missing_elements:
                            missing_elements.remove("wo CAN termination")
        if self.camera_in_title:
            if self.letter_identifier in ("R", "SC", "X"):
                # Validates that Camera exist with ID K188332H000
                result = next(
                    (
                        (title_tuple, index)
                        for index, title_tuple in enumerate(titles)
                        if "camera" in title_tuple[0].lower()
                        and title_tuple[3].upper() == "K188332H000"
                    ),
                    None,
                )
                if not result:
                    if "Camera" in missing_elements:
                        index = missing_elements.index("Camera")
                        missing_elements[index] = "Camera (K188332H000)"
                    else:
                        missing_elements.append("Camera (K188332H000)")
                # Validate the presence of 'Dataset' in missing_elements
                if "Dataset" in missing_elements:
                    # Count how many times 'Software' appears in titles
                    count_software = sum(
                        1
                        for title, _, description, _ in titles
                        if "software" in title.lower()
                    )
                    # Check if there are at least 2 'Software' elements and at least one containing 'DATASET' in the description
                    if count_software >= 2 and any(
                        (
                            "DATASET" in description.upper()
                            or "DATA SET" in description.upper()
                        )
                        for title, _, description, _ in titles
                        if "software" in title.lower()
                    ):
                        missing_elements.remove("Dataset")

                # Validate the presence of 'Software' in missing_elements
                boot_software_found = any(
                    "software" in title.lower() and "boot" in description.lower()
                    for title, _, description, _ in titles
                )
                if not boot_software_found:
                    missing_elements.append("Boot Software")
                else:
                    if "Boot Software" in missing_elements:
                        missing_elements.remove("Boot Software")
                # Check if there is an element containing "Software" in the title and matches the defined pattern, and does not have "Boot" in the description
                pattern = r"NA\d{6}"
                software_ver_found = any(
                    "software" in title.lower() and re.search(pattern, description)
                    for title, _, description, _ in titles
                    if "boot software" not in description.lower()
                )
                if not software_ver_found and not "Software" in missing_elements:
                    missing_elements.append("Software")
        if self.bracket_in_title and self.type_dropdown.value == "FLC-25 Bracket":
            # Validate the presence of 'Bracket' in missing_elements
            bracket_found = any(
                "bracket" in title.lower() and "assembly" not in title.lower()
                for title, _, description, _ in titles
            )
            if not bracket_found:
                if not "Bracket" in missing_elements:
                    missing_elements.append("Bracket")
        # Iterate over the required elements
        for element in required_elements:
            # Define the value of 'expected_value' based on the required element
            if element in ["Nut", "Screw"]:
                expected_value = 3
            else:
                expected_value = 1

            # Iterate over the titles and add the corresponding data to the DataFrame
            for title, quantity, description, _id in titles:
                qty = int(float(quantity))  # Convert quantity to a decimal number
                if (
                    self.type_dropdown.value == "FLC-25 Bracket"
                    and element == "Bracket"
                ):
                    if (
                        element.lower() in title.lower()
                        and "assembly" not in title.lower()
                    ):
                        data["ID"].append(_id)
                        data["Title"].append(title)
                        data["Description"].append(description)
                        data["Quantity"].append(qty)  # Use qty instead of quantity
                        data["Expected_value"].append(expected_value)
                else:
                    if element in title or (
                        element in ("wo CAN termination", "with CAN termination")
                        and element.lower() in description.lower()
                    ):
                        data["ID"].append(_id)
                        data["Title"].append(title)
                        data["Description"].append(description)
                        data["Quantity"].append(qty)  # Use qty instead of quantity
                        data["Expected_value"].append(expected_value)
        for i in titles:
            # Add element for CAN termination to see if it has an element:
            if (
                "wo CAN termination" in required_elements
                and "with CAN termination".lower() in i[2].lower()
            ) or (
                "with CAN termination" in required_elements
                and "wo CAN termination".lower() in i[2].lower()
            ):
                data["ID"].append(i[3])
                data["Title"].append(i[0])
                data["Description"].append(i[2])
                data["Quantity"].append(int(float(i[1])))
                data["Expected_value"].append(1)
        if self.letter_identifier == "SC" and "Screw" not in required_elements:
            for title, quantity, description, _id in titles:
                if "screw" in title.lower():
                    expected_value = 3
                    qty = int(float(quantity))
                    data["ID"].append(_id)
                    data["Title"].append(title)
                    data["Description"].append(description)
                    data["Quantity"].append(qty)
                    data["Expected_value"].append(expected_value)
                    break

        df = pd.DataFrame(data)

        def highlight_mismatched_values(val):
            color = "red" if val["Quantity"] != val["Expected_value"] else "black"
            font_weight = (
                "bold" if val["Quantity"] != val["Expected_value"] else "normal"
            )
            return [f"color: {color}; font-weight: {font_weight}" for _ in val]

        # Apply formatting to the DataFrame
        styled_df = df.style.apply(
            highlight_mismatched_values,
            axis=1,
            subset=["Title", "Description", "Quantity", "Expected_value"],
        ).set_table_styles(
            [
                {
                    "selector": "td:nth-child(2)",
                    "props": [("width", "140px")],  # 'ID' width
                },
                {
                    "selector": "td:nth-child(3)",
                    "props": [("width", "180px")],  # 'Title' width
                },
                {
                    "selector": "td:nth-child(4)",
                    "props": [("width", "450px")],  # 'Description' width
                },
                {
                    "selector": "td:nth-child(5)",
                    "props": [("width", "120px")],  # 'Quantity' width
                },
                {
                    "selector": "td:nth-child(6)",
                    "props": [("width", "120px")],  # 'Expected_value' width
                },
            ]
        )
        msg_pdf = ""
        if missing_elements:
            msg_missing = f"Missing elements: {missing_elements} for the part number {self.part_number}\n"
            print(f"\033[1m\033[48;5;208m{msg_missing}\033[0m")
            msg_pdf += msg_missing
            display(styled_df)
            return [{"msg": msg_pdf, "df": df}]

        if self.letter_identifier in ("R", "X", "SC", "B", "Cover"):
            # Rename item in required_elements
            required_elements = [
                (
                    "Radar wo CAN termination (K188333H002)"
                    if item == "wo CAN termination"
                    else item
                )
                for item in required_elements
            ]
            required_elements = [
                (
                    "Radar with CAN termination (K218450H002)"
                    if item == "with CAN termination"
                    else item
                )
                for item in required_elements
            ]
            required_elements = [
                "Camera (K188332H000)" if item == "Camera" else item
                for item in required_elements
            ]
            if self.letter_identifier == "SC":
                msg_paccar = "Please note: If the customer is Paccar, ensure that the 'Screw' element is present with a quantity of 3."
                print(f"\033[1;31m{msg_paccar}\033[0m")
                msg_pdf += msg_paccar
            msg_success_init = f"\nSUCCESSFUL:"
            msg_success = f"\nThe required elements: {required_elements}, are present for the part {self.part_number}\n"
            print(
                f"\033[1;30;48;5;34m\n{msg_success_init}\033[0m\033[1m{msg_success}\033[0m"
            )
            msg_pdf += msg_success_init + msg_success
            display(styled_df)

        return [{"msg": msg_pdf, "df": df}]

    def validate_required_docs(self, documents: dict) -> pd.DataFrame:
        """
        Validate required documents for a part number.

        Args:
            documents (dict): A dictionary containing document details.

        Returns:
            pd.DataFrame or None: DataFrame of present documents if any, else None.
        """
        msg_pdf = ""
        if not documents:
            msg_pdf = f"No documents were found in {self.part_number}"
            return [{"msg": msg_pdf, "df": pd.DataFrame()}]
        # Convert dictionary to DataFrame
        documents_df = pd.DataFrame.from_dict(documents, orient="index")

        # Define required document types and states
        if self.element_name == "Radar":
            if self.letter_identifier == "R":
                required_docs = [
                    "Assembly Drawing",
                    "Installation Drawing",
                    "Component Drawing",
                    "Service Data",
                ]  # , "Service Data"
            else:
                required_docs = [
                    "Assembly Drawing",
                    "Component Drawing",
                    "Service Data",
                ]  # , "Service Data"
        elif self.element_name == "Bracket":
            if self.type_dropdown.value in ["FLR-25 Bracket"]:
                optional_docs = [
                    "Feasibility Agreement (FeA) (DMI)",
                    "Mechanical TR Document",
                    "Electrical TR Document",
                    "Test Report",
                ]
                required_docs = [
                    "Material Specification",
                    "Supplier PPAP",
                    "Order Drawing",
                    "DVP",
                ] + optional_docs
                if self.is_am_bracket:
                    required_docs = []
                    no_doc_msg = "\nNo documents are required for this part."
                    print(f"\033[1m{no_doc_msg}\033[0m")
                    msg_pdf += no_doc_msg
            if self.type_dropdown.value in ["FLC-25 Bracket"]:
                optional_docs = ["Electrical TR Document", "Process Data Sheet"]
                required_docs = [
                    "Technical Customer Document",
                    "Assembly Drawing",
                    "Installation Drawing",
                ] + optional_docs
        elif self.element_name == "Cover" and self.type_dropdown.value in [
            "FLC-25 Cover"
        ]:
            if self.is_am_bracket:
                required_docs = ["Component Drawing"]
            else:
                required_docs = ["Installation Drawing", "Component Drawing"]
        else:  # Case Camera:
            if self.letter_identifier == "R":
                required_docs = [
                    "Assembly Drawing",
                    "Installation Drawing",
                    "Service Data",
                ]  # , "Service Data"
            else:
                required_docs = ["Assembly Drawing", "Service Data"]  # , "Service Data"
        required_states = ["Working", "Proposed", "Rejected", "Approved", "Released"]

        # Filter documents by required types
        required_documents = documents_df.query(
            "document_kind in @required_docs and lifecycle_state in @required_states"
        )

        # Initialize lists for present and missing documents
        present_docs_list = []
        missing_docs = []

        # Initialize a dictionary to track missing languages for Service Data
        service_data_missing_langs = {"US": True}

        # Check for existence of required documents
        for doc_type in required_docs:
            doc_exists = False
            for index, row in required_documents.iterrows():
                if row["document_kind"] == doc_type:
                    doc_exists = True
                    present_docs_list.append(
                        {
                            "Document": index,
                            "Document_Kind": row["document_kind"],
                            "Lifecycle_State": row["lifecycle_state"],
                        }
                    )
                    # Check languages for Service Data
                    if doc_type == "Service Data":
                        title_langs = [lang.strip() for lang in index.split(",")[1:]]
                        for lang in service_data_missing_langs:
                            if lang in title_langs:
                                service_data_missing_langs[lang] = False
                            if "EN" in title_langs:
                                service_data_missing_langs["US"] = False
            if doc_type == "Assembly Drawing":
                assembly_docs = required_documents[
                    required_documents["document_kind"] == doc_type
                ]
                if assembly_docs.empty:
                    missing_docs.extend(
                        ["Assembly Drawing Schedule", "Assembly Drawing CAD"]
                    )
                    doc_exists = True
                else:
                    if (
                        not assembly_docs.index.str.lower()
                        .str.contains("schedule")
                        .any()
                    ):
                        missing_docs.append("Assembly Drawing Schedule")
                    if assembly_docs.index.str.lower().str.contains("schedule").all():
                        missing_docs.append("Assembly Drawing CAD")
                    doc_exists = True
            if not doc_exists:
                missing_docs.append(doc_type)

        # Check for missing languages for Service Data
        if self.type_dropdown.value in ["FLC-20", "FLR-21", "FLC-25", "FLR-25"]:
            for lang, missing in service_data_missing_langs.items():
                if missing:
                    (
                        missing_docs.remove("Service Data")
                        if "Service Data" in missing_docs
                        else None
                    )
                    missing_docs.append(f"Service Data - {lang}")
        # Check if "Process Data Sheet" is in the missing list and replace it if so
        if self.type_dropdown.value in ["FLC-25 Bracket"]:
            missing_docs = [
                item if item != "Process Data Sheet" else "Process Data Sheet (BWs)"
                for item in missing_docs
            ]
        if missing_docs:
            msg_missing = f"\nThe following required documents are missing in the part {self.part_number}:\n"
            msg_pdf += msg_missing
            print(f"\033[1m\033[48;5;208m{msg_missing}\033[0m")
            for doc in missing_docs:
                doc_msg = f"- {doc}"
                print(f"\033[1;38;2;255;0;0m{doc_msg}\033[0m")
                msg_pdf += f"{doc_msg}\n"
            if self.type_dropdown.value in ["FLC-20", "FLR-21", "FLR-25", "FLC-25"]:
                if any("Service Data" in doc for doc in missing_docs):
                    serv_data_msg = f"\nPlease validate if the 'Service Data' document is required.\n"
                    print(f"\033[1m\033[38;2;255;165;0m{serv_data_msg}\033[0m")
                    msg_pdf += serv_data_msg
            if self.type_dropdown.value in ["FLR-25 Bracket", "FLC-25 Bracket"]:
                for element in optional_docs:
                    if any(element in doc for doc in missing_docs):
                        any_element_msg = f"\nPlease validate if the '{element}' document is required."
                        print(f"\033[1m\033[38;2;255;165;0m{any_element_msg}\033[0m")
                        msg_pdf += any_element_msg
        else:
            msg_success_init = f"\nSUCCESSFUL:"
            msg_success = f"\nAll required documents {required_docs} are present in the part {self.part_number}."
            print(
                f"\033[1;30;48;5;34m\n{msg_success_init}\033[0m\033[1m{msg_success}\033[0m"
            )
            msg_pdf += msg_success_init + msg_success

        if present_docs_list:
            present_doc_types = pd.DataFrame(present_docs_list).sort_values(
                by="Lifecycle_State"
            )
            msg_present_docs = f"\n\nPresent documents in {self.part_number}:\n"
            print(f"\033[1m{msg_present_docs}\033[0m")
            msg_pdf += msg_present_docs
            df_aligned = present_doc_types.style.set_table_styles(
                [dict(selector="th", props=[("text-align", "center")])]
            )
            df_aligned.set_properties(**{"text-align": "left"}).hide(axis="index")
            display(df_aligned)
            return [{"msg": msg_pdf, "df": present_doc_types}]
        else:
            return [{"msg": msg_pdf, "df": pd.DataFrame()}]

    def display_dataframe_width(self, df, title, column_widths):
        """Display DataFrame with specified column widths."""
        df_aligned = df.style.set_table_styles(
            [dict(selector="th", props=[("text-align", "center")])]
        )
        for col, width in column_widths.items():
            df_aligned.set_properties(
                subset=pd.IndexSlice[:, col], **{"width": f"{width}px"}
            )
        df_aligned.set_properties(**{"text-align": "left"}).hide(axis="index")
        display(Markdown(f"<span style='font-size:14px'>{title}</span>"))
        display(df_aligned)

    def validate_aftermarket_product(self, progress_bar) -> list:
        """Validate aftermarket product documents and display them."""
        if self.type_dropdown.value in [
            "FLC-25 Bracket",
            "FLR-25 Bracket",
            "FLC-25 Cover",
        ]:
            if self.is_am_bracket:
                docs_aftermarket_for = self.utils.documents_aftermarket_product_for(
                    self.driver
                )
            else:
                docs_aftermarket_for = None
                docs_aftermarket = self.utils.documents_aftermarket_product(self.driver)
        else:
            docs_aftermarket = self.utils.documents_aftermarket_product(self.driver)
        progress_bar.update(15)

        required_doc_aftermarket = [self.element_name]

        lifecycle_state = ["Working", "Proposed", "Rejected", "Approved", "Released"]
        if not self.is_am_bracket:
            if docs_aftermarket:
                aftermarket_df = pd.DataFrame.from_dict(
                    docs_aftermarket, orient="index"
                )
                aftermarket_df = aftermarket_df[
                    (aftermarket_df["Lifecycle_State"].isin(lifecycle_state))
                    & (
                        aftermarket_df["Title_ID_value"].str.contains(
                            "|".join(required_doc_aftermarket)
                        )
                    )
                ]
                aftermarket_df = aftermarket_df.sort_values(by="Lifecycle_State")
                missing_aftermarket_elements = [
                    elem
                    for elem in required_doc_aftermarket
                    if not any(
                        elem.lower() in title.lower()
                        for title in aftermarket_df["Title_ID_value"]
                    )
                ]
            else:
                aftermarket_df = pd.DataFrame(
                    columns=["Document", "Title_ID_value", "Lifecycle_State"]
                )
                missing_aftermarket_elements = required_doc_aftermarket
        else:
            if docs_aftermarket_for:
                aftermarket_df = pd.DataFrame.from_dict(
                    docs_aftermarket_for, orient="index"
                )
                aftermarket_df = aftermarket_df[
                    (aftermarket_df["Lifecycle_State"].isin(lifecycle_state))
                    & (
                        aftermarket_df["Title_ID_value"].str.contains(
                            "|".join(required_doc_aftermarket)
                        )
                    )
                ]
                aftermarket_df = aftermarket_df.sort_values(by="Lifecycle_State")
                missing_aftermarket_elements = [
                    elem
                    for elem in required_doc_aftermarket
                    if not any(
                        elem.lower() in title.lower()
                        for title in aftermarket_df["Title_ID_value"]
                    )
                ]
            else:
                aftermarket_df = pd.DataFrame(
                    columns=["Document", "Title_ID_value", "Lifecycle_State"]
                )
                missing_aftermarket_elements = required_doc_aftermarket
        msg_pdf_am = ""
        if missing_aftermarket_elements:
            msg_missing_am = (
                "The following required documents in Aftermarket Product are missing:\n"
            )
            print(f"\033[1m\033[48;5;208m{msg_missing_am}\033[0m")
            msg_pdf_am += msg_missing_am
            for doc in missing_aftermarket_elements:
                doc_msg_am = f"- {doc}"
                print(f"\033[1;38;2;255;0;0m{doc_msg_am}\033[0m")
                msg_pdf_am += f"{doc_msg_am}\n"
            print("\n")
        # Display DataFrames with specified widths
        column_widths = {"Document": 180, "Title_ID_value": 200, "Lifecycle_State": 100}
        if not aftermarket_df.empty:
            print(f"\033[1;30;48;5;34m\nSUCCESSFUL: \n\033[0m")
            msg_pdf_am += (
                f"Results for AfterMarket Product in part {self.part_number}:\n"
            )
            self.display_dataframe_width(
                aftermarket_df, "**AfterMarket Product:**", column_widths
            )
            print("\n")
        return [{"msg": msg_pdf_am, "df": aftermarket_df}]
