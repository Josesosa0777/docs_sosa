import logging
import os
import re
import time
from typing import Dict, List, Optional, Tuple

import pandas as pd
from IPython.display import display, Markdown
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.remote.webdriver import WebDriver

from enums import ElementType, LetterIdentifier, LifeCycleStatus, DocumentType, BomScheduleKey

# Constants
TIME_SLEEP_BASE = 3
TIME_SLEEP_ADDITIONAL = 2
EXTRA_TIME_SCALE = 0.5
DEFAULT_TIMEOUT = 10

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class Legacy:
    """Interact with legacy parts."""

    REQUIRED_DOCS = [DocumentType.ASSEMBLY_DRAWING, DocumentType.INSTALLATION_DRAWING, DocumentType.SERVICE_DATA]
    REQUIRED_STATES = [LifeCycleStatus.WORKING, LifeCycleStatus.PROPOSED, LifeCycleStatus.REJECTED,
                        LifeCycleStatus.APPROVED, LifeCycleStatus.RELEASED]

    def __init__(self, utils, extra_time=0):
        self.extra_time_sec = int(extra_time)
        self.utils = utils

    def init_schedule(self, driver: WebDriver, schedule_dropdown: str, chbx_schedule: str) -> None:
        """Initialize schedule parameters."""
        self.schedule_dropdown = schedule_dropdown
        self.driver = driver
        self.chbx_schedule = chbx_schedule

    def initialize_main_params(
        self,
        part_number: str,
        title_id_value: str,
        driver: WebDriver,
        child_div_path: str,
        letter_identifier: str,
        chbx_bracket: str,
        chbx_resistor: str,
        type_dropdown: str,
    ) -> None:
        """Initialize main parameters and check for radar or camera in title."""
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
        if not (self.radar_in_title or self.camera_in_title):
            raise ValueError("No specific required elements defined for camera or radar.")
        self.element_name = "Radar" if self.radar_in_title else "Camera"

    # --- INI File Section ---
    def select_doc_dataset(self, dataset_idx: int) -> None:
        """Select the Document Dataset and check for INI file."""
        try:
            dataset_locator = By.XPATH, f"//*[@id='occTreeTable_row{dataset_idx}_col1']"
            self.utils.click_element(self.driver, dataset_locator)
            time.sleep(TIME_SLEEP_BASE + self.extra_time_sec)

            self.utils.click_element(self.driver, By.XPATH, "//*[@id='main-view']/div/div[2]/div[1]/div/div/div/div/div[3]/div/div/div[1]/div/div/main/div[3]/div/div/div[1]/div/ul/li[4]/a")
            time.sleep(12 + self.extra_time_sec)
            
            self.utils.click_element(self.driver, By.XPATH, "//div[contains(@class, 'sw-row sw-sectionTitleContainer')]//div[@title='Documents']")
            time.sleep(TIME_SLEEP_BASE + self.extra_time_sec)

            divs_docs = self.driver.find_elements(By.XPATH, "//*[@id='ObjectSet_6_Provider']/div[2]/div[2]/div[2]/div/div")
            has_ini_file = False

            for index, div in enumerate(divs_docs, start=1):
                    title = self.utils.wait_for_element(self.driver, By.XPATH, f"//*[@id='ObjectSet_6_Provider']/div[2]/div[3]/div[2]/div/div[{index}]/div[1]/div", timeout=DEFAULT_TIMEOUT).get_attribute("title")
                    if title.lower() == "software":
                        life_cycle = self.utils.wait_for_element(self.driver, By.XPATH, f"//*[@id='ObjectSet_6_Provider']/div[2]/div[3]/div[2]/div/div[{index}]/div[3]/div", timeout=DEFAULT_TIMEOUT).get_attribute("title")
                        try:
                            if LifeCycleStatus(life_cycle) in LifeCycleStatus:
                                if LifeCycleStatus(life_cycle) in [LifeCycleStatus.WORKING, LifeCycleStatus.PROPOSED, LifeCycleStatus.REJECTED, LifeCycleStatus.APPROVED, LifeCycleStatus.RELEASED]:
                                    div.click()
                                    self.utils.click_element(self.driver, By.XPATH, f"//*[@id='ObjectSet_6_Provider_row{index+1}_col2']/div/div[2]")
                                    time.sleep(TIME_SLEEP_ADDITIONAL + self.extra_time_sec)
                                    has_ini_file = True
                        except ValueError:
                            logger.warning(f"Unrecognized life cycle status: {life_cycle}")
            if not has_ini_file:
                    logger.info("INI file was not located.")

        except Exception as e:
            logger.error(f"An error occurred while selecting the document dataset: {e}")

    def expand_elements_get_titles(self) -> List[Tuple[str, str, str, str]]:
        """Expand elements and get titles."""
        time.sleep(EXTRA_TIME_SCALE + EXTRA_TIME_SCALE * self.extra_time_sec)
        separator = self.driver.find_element(By.XPATH, "//*[@id='main-view']/div/div[2]/div[1]/div/div/div/div/div[3]/div/div/div[1]/div/div/main/div[2]")
        actions = ActionChains(self.driver)

        element_location = self.driver.execute_script("return arguments[0].getBoundingClientRect();", separator)
        viewport_width = self.driver.execute_script("return window.innerWidth;")
        move_x = min(250, viewport_width - element_location["left"] - element_location["width"])

        actions.click_and_hold(separator).move_by_offset(move_x, 0).release().perform()
        div_elements = self.driver.find_elements(By.XPATH, f"{self.child_div_path}[@aria-level='2']")
        titles = self.utils.get_titles(self.driver, div_elements, 2)

        if self.letter_identifier in (LetterIdentifier.N, LetterIdentifier.SC):
            idx = next((index for index, (title, _, _, _) in enumerate(titles) if self.element_name.value.lower() in title.lower()), None)
            if idx is not None:
                self.utils.click_element(self.driver, By.XPATH, f"//main/div[1]/div/div[2]/div/aw-splm-table/div/div[2]/div[2]/div[2]/div/div[{idx+2}]/div/div/div/div[1]")
                time.sleep(TIME_SLEEP_ADDITIONAL + EXTRA_TIME_SCALE * self.extra_time_sec)
                
                elements_x = self.driver.find_elements(By.XPATH, f"{self.child_div_path}[@aria-level='3']")
                if elements_x:
                    titles.pop(idx)
                time.sleep(TIME_SLEEP_ADDITIONAL + EXTRA_TIME_SCALE * self.extra_time_sec)
                child_titles = self.utils.get_titles(self.driver, elements_x, idx + 3)
                titles.extend(child_titles)
        
        actions.click_and_hold(separator).move_by_offset(-move_x, 0).release().perform()
        return titles

    def get_required_elements(self, element_name: ElementType) -> List[str]:
        """Get required elements based on element name and letter identifier."""
        elements_dict = {
            ElementType.RADAR: {
                LetterIdentifier.X: ["Radar", "Software", "Dataset", "Label"],
                LetterIdentifier.R: ["Adjustor", "Screw", "Radar", "Software", "Bracket", "Dataset", "Label"],
                LetterIdentifier.N: ["Radar", "Literature", "Adjusting", "Software", "Dataset", "Label"],
                LetterIdentifier.SC: ["Radar", "Literature", "Adjusting", "Software", "Dataset", "Label"],
            },
            ElementType.CAMERA: {
                LetterIdentifier.X: ["Dataset", "Label", "Camera", "Software"],
                LetterIdentifier.R: ["Dataset", "Label", "Camera", "Software"],
                LetterIdentifier.N: ["Dataset", "Label", "Camera", "Software"],
                LetterIdentifier.SC: ["Dataset", "Label", "Camera", "Software"],
            }
        }
        if isinstance(self.letter_identifier, str):
            try:
                lettr = LetterIdentifier[self.letter_identifier]
            except KeyError:
                print(f"Invalid LetterIdentifier string: {self.letter_identifier}")
                return []
        return elements_dict.get(element_name, {}).get(lettr, [])

    def get_bom_in_content(self) -> List[Dict[str, pd.DataFrame]]:
        """Retrieve BOM content, validate required elements, and generate DataFrame with results."""
        # Expand elements and get titles
        titles = self.expand_elements_get_titles()
        time.sleep(1 + EXTRA_TIME_SCALE * self.extra_time_sec)
        required_elements = self.get_required_elements(ElementType.RADAR if self.radar_in_title else ElementType.CAMERA)

        missing_elements = [element for element in required_elements if not any(element in title for title, _, _, _ in titles)]

        # Prepare data for DataFrame
        data = {
            "ID": [],
            "Title": [],
            "Description": [],
            "Quantity": [],
            "Expected_value": [],
        }

        if self.radar_in_title and self.letter_identifier == LetterIdentifier.R.value:
            self._handle_radar_missing_elements(missing_elements, titles)

        for element in required_elements:
            expected_value = self._get_expected_value(element)
            for title, quantity, description, _id in titles:
                qty = int(float(quantity))  # Convert quantity to a decimal number
                if element in title:
                    data["ID"].append(_id)
                    data["Title"].append(title)
                    data["Description"].append(description)
                    data["Quantity"].append(qty)
                    data["Expected_value"].append(expected_value)

        df = pd.DataFrame(data)

        styled_df = self._style_comparison_result(df)

        return self._generate_report(missing_elements, required_elements, titles, styled_df, df)

    def _style_comparison_result(self, comparison_result: pd.DataFrame) -> pd.DataFrame.style:
        """
        Apply styling to the comparison result DataFrame.

        Args:
            comparison_result (pd.DataFrame): DataFrame with comparison results.

        Returns:
            pd.DataFrame.style: Styled DataFrame.
        """
        def highlight_mismatched_values(row):
            color = "red" if row["Expected_value"] != row["Quantity"] else "black"
            font_weight = "bold" if row["Expected_value"] != row["Quantity"] else "normal"
            return [f"color: {color}; font-weight: {font_weight}" for _ in row]

        return comparison_result.style.apply(
            highlight_mismatched_values,
            axis=1,
            subset=["Title", "Description", "Quantity", "Expected_value"]
        ).set_table_styles(
            [
                {"selector": "th", "props": [("text-align", "left"), ("font-weight", "bold")]},
                {"selector": "td:nth-child(1)", "props": [("width", "120px"), ("text-align", "left"), ("word-break", "break-word")]},
                {"selector": "td:nth-child(2)", "props": [("width", "250px"), ("text-align", "left"), ("word-break", "break-word")]},
                {"selector": "td:nth-child(3)", "props": [("width", "400px"), ("text-align", "left"), ("word-break", "break-word")]},
                {"selector": "td:nth-child(4)", "props": [("width", "120px"), ("text-align", "right")]},
                {"selector": "td:nth-child(5)", "props": [("width", "120px"), ("text-align", "right")]},
            ]
        )

    def _handle_radar_missing_elements(self, missing_elements: List[str], titles: List[Tuple[str, str, str, str]]) -> None:
        """Handle specific validation for radar missing elements."""
        if "Dataset" in missing_elements:
            count_software = sum(
                1 for title, _, _, _ in titles if "software" in title.lower()
            )
            if count_software >= 2 and any(
                "DATASET" in description.upper() or "DATA SET" in description.upper()
                for title, _, description, _ in titles
                if "software" in title.lower()
            ):
                missing_elements.remove("Dataset")

        if "Software" not in missing_elements:
            pattern = r"BX\d{6}"
            if not any(
                re.search(pattern, description)
                for title, _, description, _ in titles
                if "software" in title.lower()
            ):
                missing_elements.append("Software")
                logger.warning("No Software with 'BX' followed by 6 digits")

    def _generate_report(self, missing_elements: List[str], required_elements: List[str], titles: List[Tuple[str, str, str, str]], styled_df: pd.DataFrame.style, df: pd.DataFrame) -> List[Dict[str, pd.DataFrame]]:
        """Generate report based on missing elements and DataFrame."""
        msg_pdf = ""
        if missing_elements:
            msg_missing = f"Missing elements: {missing_elements} for the part number {self.part_number}\n"
            print(f"\033[1m\033[48;5;208m{msg_missing}\033[0m")
            msg_pdf += msg_missing
            if self.radar_in_title and self.letter_identifier == LetterIdentifier.R.value:
                count_label = sum(
                    1 for title, _, _, _ in titles if "label" in title.lower()
                )
                if count_label != 2:
                    msg_label = f"\nConsider that there should be 2 'Label' elements, found {count_label}"
                    print(msg_label)
                    msg_pdf += msg_label
            display(styled_df)
            return [{"msg": msg_pdf, "df": df}]

        if self.letter_identifier in (LetterIdentifier.N.value, LetterIdentifier.SC.value, LetterIdentifier.X.value, LetterIdentifier.R.value):
            msg_success_init = f"\nSUCCESSFUL:"
            msg_success = f"\nThe required elements: {required_elements}, are present for the part {self.part_number}\n"
            print(f"\033[1;30;48;5;34m\n{msg_success_init}\033[0m\033[1m{msg_success}\033[0m")
            msg_pdf += msg_success_init + msg_success
            display(styled_df)
            if self.radar_in_title and self.letter_identifier == LetterIdentifier.R.value:
                count_label = sum(
                    1 for title, _, _, _ in titles if "label" in title.lower()
                )
                if count_label != 2:
                    msg_label = f"\nConsider that there should be 2 'Label' elements, found {count_label}"
                    print(msg_label)
                    msg_pdf += msg_label
        return [{"msg": msg_pdf, "df": df}]

    def _get_expected_value(self, element: str) -> int:
        """Return the expected quantity value for a given element."""
        expected_values = {
            "Adjustor": 3,
            "Screw": 6
        }
        return expected_values.get(element, 1)

    def validate_required_docs(self, documents: Dict) -> List[Dict]:
        """
        Validate required documents for a part number.

        Args:
            documents (dict): A dictionary containing document details.

        Returns:
            List[Dict]: A list containing message and DataFrame of present documents if any, else empty list.
        """
        if not documents:
            msg = f"No documents were found for part number {self.part_number}"
            return [{"msg": msg, "df": pd.DataFrame()}]

        # Convert dictionary to DataFrame
        documents_df = pd.DataFrame.from_dict(documents, orient="index")

        # Filter documents by required types and states
        required_docs_list = [doc.value for doc in Legacy.REQUIRED_DOCS]
        required_states_list = [doc.value for doc in Legacy.REQUIRED_STATES]
        required_documents = documents_df[
            documents_df['document_kind'].isin(required_docs_list) &
            documents_df['lifecycle_state'].isin(required_states_list)
        ]

        present_docs_list, missing_docs = self._check_documents(required_documents)

        msg_pdf = self._generate_message(present_docs_list, missing_docs, required_docs_list)

        return [{"msg": msg_pdf, "df": pd.DataFrame(present_docs_list)}]

    def get_required_states(self):
        """Get the string values of required lifecycle states."""
        return [state.value for state in Legacy.REQUIRED_STATES]
        
    def _check_documents(self, required_documents: pd.DataFrame) -> Tuple[List[Dict[str, any]], List[str]]:
        """
        Check for existence of required documents and handle missing ones.

        Args:
            required_documents (pd.DataFrame): Filtered DataFrame containing required documents.

        Returns:
            Tuple[List[Dict[str, any]], List[str]]: List of present documents and list of missing documents.
        """
        present_docs_list = []
        missing_docs = []
        service_data_missing_langs = {"ES": True, "US": True, "FR": True}

        for doc_type in Legacy.REQUIRED_DOCS:
            doc_exists = [self._check_document_existence(doc_type, required_documents, service_data_missing_langs, present_docs_list)]
            
            if doc_type == DocumentType.ASSEMBLY_DRAWING:
                self._check_assembly_drawings(required_documents, present_docs_list, missing_docs, doc_exists)

            if not doc_exists[0]:
                missing_docs.append(doc_type.value)

        self._handle_service_data_missing_languages(service_data_missing_langs, missing_docs)

        return present_docs_list, missing_docs

    def _check_document_existence(self, doc_type: DocumentType, required_documents: pd.DataFrame, 
                                   service_data_missing_langs: Dict[str, bool], present_docs_list: List[Dict]) -> bool:
        """
        Check if a specific document type exists and update lists accordingly.

        Args:
            doc_type (DocumentType): The document type to check.
            required_documents (pd.DataFrame): DataFrame of required documents.
            service_data_missing_langs (Dict[str, bool]): Dictionary to track missing languages.
            present_docs_list (List[Dict]): List to store present documents.

        Returns:
            bool: True if the document type exists, False otherwise.
        """
        doc_exists = False
        for index, row in required_documents.iterrows():
            if row["document_kind"] == doc_type.value:
                doc_exists = True
                present_docs_list.append(
                    {
                        "Document": index,
                        "Document_Kind": row["document_kind"],
                        "Lifecycle_State": row["lifecycle_state"],
                    }
                )
                if doc_type == DocumentType.SERVICE_DATA:
                    self._update_service_data_languages(index, service_data_missing_langs)

        return doc_exists
    
    def _update_service_data_languages(self, index: str, service_data_missing_langs: Dict[str, bool]) -> None:
        """
        Update the missing languages dictionary based on the document index.

        Args:
            index (str): Document index.
            service_data_missing_langs (Dict[str, bool]): Dictionary to track missing languages.
        """
        title_langs = [lang.strip() for lang in index.split(",")[1:]]
        for lang in service_data_missing_langs:
            if lang in title_langs:
                service_data_missing_langs[lang] = False
            if "EN" in title_langs:
                service_data_missing_langs["US"] = False

    def _check_assembly_drawings(self, required_documents: pd.DataFrame, present_docs_list: List[dict], missing_docs: List[str], doc_exists: List[bool]) -> None:
        """
        Check for the presence of Assembly Drawings and handle missing types.

        Args:
            required_documents (pd.DataFrame): Filtered DataFrame of required documents.
            present_docs_list (List[dict]): List to store present documents.
            missing_docs (List[str]): List to store missing documents.
        """
        assembly_docs = required_documents[required_documents["document_kind"] == DocumentType.ASSEMBLY_DRAWING.value]
        if assembly_docs.empty:
            missing_docs.extend(["Assembly Drawing Schedule", "Assembly Drawing CAD"])
            doc_exists[0] = True
        else:
            if not assembly_docs.index.str.contains("schedule", case=False).any():
                missing_docs.append("Assembly Drawing Schedule")
            if assembly_docs.index.str.contains("schedule", case=False).all():
                missing_docs.append("Assembly Drawing CAD")
            doc_exists[0] = True

    def _handle_service_data_missing_languages(self, service_data_missing_langs: dict, missing_docs: List[str]) -> None:
        """
        Handle missing languages for Service Data documents.

        Args:
            service_data_missing_langs (dict): Dictionary to track missing languages.
            missing_docs (List[str]): List to store missing documents.
        """
        for lang, missing in service_data_missing_langs.items():
            if missing:
                if DocumentType.SERVICE_DATA.value in missing_docs:
                    missing_docs.remove(DocumentType.SERVICE_DATA.value)
                missing_docs.append(f"{DocumentType.SERVICE_DATA.value} - {lang}")

    def _generate_message(self, present_docs: List[Dict], missing_docs: List[str], required_docs_list: List[str]) -> str:
        """
        Generate message for present and missing documents.

        Args:
            present_docs (List[Dict]): List of present documents.
            missing_docs (List[str]): List of missing documents.

        Returns:
            str: Message for the documents.
        """
        msg_pdf = ""
        if missing_docs:
            msg_missing = f"\nThe following required documents are missing for part number {self.part_number}:\n"
            print(f"\033[1m\033[48;5;208m{msg_missing}\033[0m")
            msg_pdf += msg_missing
            for doc in missing_docs:
                doc_msg = f"- {doc}"
                print(f"\033[1;38;2;255;0;0m{doc_msg}\033[0m")
                msg_pdf += f"{doc_msg}\n"
        else:
            msg_success_init = "\nSUCCESSFUL:"
            msg_success = f"\nAll required documents {required_docs_list} are present for part number {self.part_number}."
            print(f"\033[1;30;48;5;34m\n{msg_success_init}\033[0m\033[1m{msg_success}\033[0m")
            msg_pdf += f"{msg_success_init}\n{msg_success}"

        if present_docs:
            present_docs_df = pd.DataFrame(present_docs).sort_values(
                by="Lifecycle_State"
            )
            msg_present_docs = f"\n\nPresent documents for part number {self.part_number}:\n"
            print(f"\033[1m{msg_present_docs}\033[0m")
            msg_pdf += msg_present_docs
            df_aligned = present_docs_df.style.set_table_styles(
                [dict(selector="th", props=[("text-align", "center")])]
            ).set_properties(**{"text-align": "left"}).hide(axis="index")
            display(df_aligned)

        return msg_pdf

    def download_excel_file(self) -> Optional[str]:
        """
        Download an Excel file based on specific conditions and return its path.
        Returns:
            Optional[str]: The path to the downloaded file if successful, otherwise `None`.
        """
        time.sleep(TIME_SLEEP_BASE * 4 + self.extra_time_sec)
        file_path = None
        div_file_elements = self.driver.find_elements(
            By.XPATH, "//*[@id='ObjectSet_2_Provider']/div[2]/div[2]/div[2]/div/div"
        )

        for i, _ in enumerate(div_file_elements, start=2):
            headers_idx = self.utils.get_headers_idx(
                self.driver,
                path_headers="//*[@id='ObjectSet_2_Provider']/div[2]/div[3]/div[1]/div",
                headers=["File Name", "Type"],
            )
            file_title = self.utils.wait_for_element(
                self.driver,
                (By.XPATH, f"//*[@id='ObjectSet_2_Provider_row{i}_col{headers_idx['Type']+3}']/div"),
                timeout=DEFAULT_TIMEOUT,
            ).get_attribute("title")

            file_name = self._get_file_name(headers_idx, i, file_title)
            if file_name:
                file_path = self._handle_existing_file(file_name)
                self._click_file_element(i)
                break

        return file_path
    
    def _get_file_name(self, headers_idx: Dict[str, int], i: int, file_title: str) -> Optional[str]:
        """
        Get the file name based on file title and type.

        Args:
            headers_idx (Dict[str, int]): Indexes of headers for file name and type.
            i (int): Index of the file element.
            file_title (str): Title of the file element.

        Returns:
            Optional[str]: The file name if it matches the conditions, otherwise `None`.
        """
        if (self.chbx_schedule.value and ("MS ExcelX" in file_title or "MS Excel" in file_title)) or \
            (not self.chbx_schedule.value and ((self.type_dropdown.value in ["FLC-25", "FLR-25"] and "Zip" in file_title) or \
            (self.type_dropdown.value in ["FLC-20", "FLR-21"] and ("MS ExcelX" in file_title or "MS Excel" in file_title)))):
            file_name = self.utils.wait_for_element(
                self.driver,
                (By.XPATH, f"//*[@id='ObjectSet_2_Provider_row{i}_col{headers_idx['File Name']+3}']/div"),
                timeout=DEFAULT_TIMEOUT
            ).get_attribute("title")
            return file_name
        return None

    def _handle_existing_file(self, file_name: str) -> str:
        """
        Handle an existing file by removing it before re-downloading.

        Args:
            file_name (str): The name of the file to be handled.

        Returns:
            str: The path to the file after handling any existing file.
        """
        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
        file_path = os.path.join(downloads_path, file_name)
        if os.path.exists(file_path):
            os.remove(file_path)
            while os.path.exists(file_path):
                time.sleep(TIME_SLEEP_ADDITIONAL + self.extra_time_sec)
        return file_path
    
    def _click_file_element(self, i: int):
        """
        Click the file element to initiate the download.

        Args:
            i (int): Index of the file element to be clicked.
        """
        self.utils.click_element(
            self.driver,
            (By.XPATH, f'//*[@id="ObjectSet_2_Provider"]/div[2]/div[2]/div[2]/div/div[{i-1}]')
        )
        self.utils.click_element(
            self.driver,
            (By.XPATH, f'//*[@id="ObjectSet_2_Provider"]/div[2]/div[2]/div[2]/div/div[{i-1}]/div[2]/div[1]/div[2]')
        )
        time.sleep(5 + self.extra_time_sec)

    def documents_spare_kitpart(self) -> Dict[int, Dict[str, str]]:
        """
        Select and return spare kit-part documents based on specified conditions.

        Returns:
            Dict[int, Dict[str, str]]: A dictionary where keys are integer indices and values are dictionaries
            containing document details, including "Document", "Title_ID_value", and "Lifecycle_State".
        """
        time.sleep(1 + 0.25 * self.extra_time_sec)
        documents = {}
        self.utils.click_element(
            self.driver,
            (By.XPATH, f"//*[@id='main-view']//main/div/details[7]/summary/div/div"),
        )
        time.sleep(1 + 0.25 * self.extra_time_sec)
        div_elements = self.driver.find_elements(
            By.XPATH, "//*[@id='ObjectSet_19_Provider']/div[2]/div[2]/div[2]/div/div"
        )
        idx = 0
        for idx, div in enumerate(div_elements, start=1):
            time.sleep(1 + self.extra_time_sec)
            aria_rowindex = div.get_attribute("aria-rowindex")
            headers_idx = self.utils.get_headers_idx(
                self.driver,
                path_headers="//*[@id='ObjectSet_19_Provider']/div[2]/div[3]/div[1]/div",
                headers=["Title ID Value", "Lifecycle State"]
            )
            document_data = {
                "Document": self._get_element_title(aria_rowindex, headers_idx, "Document"),
                "Title_ID_value": self._get_element_title(aria_rowindex, headers_idx, "Title ID Value"),
                "Lifecycle_State": self._get_element_title(aria_rowindex, headers_idx, "Lifecycle State")
            }
            documents[idx] = document_data
        return documents

    def _get_element_title(self, aria_rowindex: str, headers_idx: Dict[str, int], column_name: str) -> str:
        """
        Get the title attribute of an element from a specific column.

        Args:
            aria_rowindex (str): The aria-rowindex of the element.
            headers_idx (Dict[str, int]): Indexes of the headers for different columns.
            column_name (str): The name of the column to retrieve the title from.

        Returns:
            str: The title attribute of the specified element.
        """
        path = f"//*[@id='ObjectSet_19_Provider_row{aria_rowindex}_col{headers_idx[column_name]+3}']/div"
        if column_name == "Document":
            path = f"//*[@id='ObjectSet_19_Provider_row{aria_rowindex}_col2']/div"
        return self.utils.wait_for_element(
            self.driver,
            (By.XPATH, path),
            timeout=DEFAULT_TIMEOUT
        ).get_attribute("title")

    def display_dataframe_width(self,  df: pd.DataFrame, title: str, column_widths: Dict[str, int]):
        """
        Display a DataFrame with specified column widths and a title.

        Args:
            df (pd.DataFrame): The DataFrame to be displayed.
            title (str): The title to be displayed above the DataFrame.
            column_widths (Dict[str, int]): A dictionary mapping column names to their desired widths (in pixels).
        """
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

    def validate_aftermarket_product(self, progress_bar) -> List[Dict[str, pd.DataFrame]]:
        """
        Validate and display aftermarket product and spare kit-part documents.

        This method validates the required documents for aftermarket products and spare kit-parts. It updates the progress bar during the process and displays the results if the DataFrames are not empty.

        Args:
            progress_bar (ProgressBar): An instance of a progress bar for tracking validation progress.

        Raises:
            ValueError: If neither radar nor camera is specified in the title ID value, indicating that no specific required elements are defined.

        Returns:
            List[Dict[str, pd.DataFrame]]: A list containing two dictionaries. Each dictionary has the following keys:
                - "msg": A string message summarizing the validation results.
                - "df": A pandas DataFrame with the validation results.
        """
        docs_aftermarket = self.utils.documents_aftermarket_product(self.driver)
        progress_bar.update(10)
        docs_spare_kitpart = self.documents_spare_kitpart()
        progress_bar.update(5)

        radar_in_title = ElementType.RADAR.value.lower() in self.title_id_value.lower()
        camera_in_title = ElementType.CAMERA.value.lower() in self.title_id_value.lower()
        if not (radar_in_title or camera_in_title):
            raise ValueError("No specific required elements defined for camera or radar.")

        element_name = ElementType.RADAR.value if radar_in_title else ElementType.CAMERA.value
        required_doc_aftermarket = [element_name]
        required_doc_spare = ["Bracket"] if camera_in_title else ["Bracket", "Adjusting Device"]

        lifecycle_state = self.get_required_states()

        aftermarket_df, missing_aftermarket_elements = self._validate_docs(
            docs_aftermarket, required_doc_aftermarket, lifecycle_state
        )
        spare_kitpart_df, missing_spare_elements = self._validate_docs(
            docs_spare_kitpart, required_doc_spare, lifecycle_state
        )

        msg_pdf_am, msg_pdf_spare = self._generate_messages(
            missing_aftermarket_elements, missing_spare_elements
        )

        if not aftermarket_df.empty and not spare_kitpart_df.empty:
            print(f"\033[1;30;48;5;34m\nSUCCESSFUL:\n\033[0m")

        # Display DataFrames with specified widths
        column_widths = {"Document": 180, "Title_ID_value": 200, "Lifecycle_State": 100}
        if not aftermarket_df.empty:
            msg_pdf_am += f"Results for AfterMarket Product in part {self.part_number}:\n"
            self.display_dataframe_width(aftermarket_df, "**AfterMarket Product:**", column_widths)
        if not spare_kitpart_df.empty:
            msg_pdf_spare += f"\nResults for Spare Kit-Part in part {self.part_number}:\n"
            self.display_dataframe_width(spare_kitpart_df, "**Spare Kit-Part Documents:**", column_widths)

        return [{"msg": msg_pdf_am, "df": aftermarket_df}, {"msg": msg_pdf_spare, "df": spare_kitpart_df}]

    def _validate_docs(self, docs: Dict[int, Dict[str, str]], required_docs: List[str], required_states: List[str]) -> Tuple[pd.DataFrame, List[str]]:
        """
        Validate documents and filter based on required docs and states.

        Args:
            docs (dict): Dictionary of documents.
            required_docs (list): List of required document types.
            required_states (list): List of required lifecycle states.

        Returns:
            Tuple: DataFrame of valid documents and list of missing required documents.
        """
        docs_df = pd.DataFrame.from_dict(docs, orient="index")
        filtered_df = docs_df[
            docs_df["Title_ID_value"].str.contains("|".join(required_docs)) &
            docs_df["Lifecycle_State"].isin(required_states)
        ].sort_values(by="Lifecycle_State")

        missing_elements = [doc for doc in required_docs if not any(doc.lower() in title.lower() for title in filtered_df["Title_ID_value"])]

        return filtered_df, missing_elements

    def _generate_messages(self, missing_aftermarket_elements: List[str], missing_spare_elements: List[str]) -> Tuple[str, str]:
        """
        Generate messages for missing documents.

        Args:
            missing_aftermarket_elements (list): List of missing aftermarket elements.
            missing_spare_elements (list): List of missing spare kit-part elements.

        Returns:
            Tuple: Messages for missing aftermarket and spare kit-part elements.
        """
        msg_pdf_am = ""
        msg_pdf_spare = ""

        if missing_aftermarket_elements:
            msg_missing_am = f"The following required documents in Aftermarket Product are missing for the part {self.part_number}:\n"
            print(f"\033[1m\033[48;5;208m{msg_missing_am}\033[0m")
            msg_pdf_am += msg_missing_am
            for doc in missing_aftermarket_elements:
                doc_msg_am = f"- {doc}"
                print(f"\033[1;38;2;255;0;0m{doc_msg_am}\033[0m")
                msg_pdf_am += f"{doc_msg_am}\n"
            print("\n")

        if missing_spare_elements:
            msg_missing_spare = f"The following required documents in Spare Kit-Part are missing for the part {self.part_number}:\n"
            print(f"\033[1m\033[48;5;208m{msg_missing_spare}\033[0m")
            msg_pdf_spare += msg_missing_spare
            for doc in missing_spare_elements:
                doc_msg_spare = f"- {doc}"
                print(f"\033[1;38;2;255;0;0m{doc_msg_spare}\033[0m")
                msg_pdf_spare += f"{doc_msg_spare}\n"
            print("\n")

        return msg_pdf_am, msg_pdf_spare

    def get_titles_schedule_in_content(self) -> Tuple[List[str], List[str]]:
        """
        Get titles in Content tab.

        Returns:
            Tuple[List[str], List[str]]:
                - A list of detailed titles obtained from `expand_elements_get_titles`.
                - A list of main titles extracted from the main elements on the Content tab.
        """
        titles = self.expand_elements_get_titles()
        main_elements = self.driver.find_elements(
            By.XPATH, f"{self.child_div_path}[@aria-level='1']"
        )
        main_title = self.utils.get_titles(self.driver, main_elements, 1)
        return titles, main_title

    def get_bom_schedule_in_content(self, schedule_data: Dict[str, str]) -> Dict[str, pd.DataFrame]:
        """
        Compares schedule data with BOM titles to determine if values are correct.

        Args:
            schedule_data (Dict[str, str]): A dictionary containing the schedule data with parameter values.

        Returns:
            Dict[str, pd.DataFrame]: A dictionary containing a message and a DataFrame with comparison results.
        """
        titles, main_title = self.get_titles_schedule_in_content()

        # Convert keys to enum for better readability and maintainability
        data_keys = [key.value for key in BomScheduleKey]
        
        rows = []
        for key, value in schedule_data.items():
            if not key in data_keys:
                continue
            is_in_bom = self._check_key_in_bom(key, value, titles, main_title)
            correct_value = "Correct" if is_in_bom else "Incorrect"

            rows.append({
                "Parameter": key,
                "Schedule Value": value,
                "Comparison vs BOM": correct_value,
            })

        df = pd.DataFrame(rows)
        df = self._post_process_dataframe(df)

        msg_pdf = f"\nResults for Part {schedule_data.get(BomScheduleKey.PART_NUMBER.value, 'Unknown')}:\n"
        print(f"\033[1m{msg_pdf}\033[0m")
        styled_df = self.style_schedule_result(df)
        display(styled_df)
        print("\n")
        
        return {"msg": msg_pdf, "df": df}
    
    def _check_key_in_bom(self, key: str, value: str, titles: List[Tuple], main_title: List[Tuple]) -> bool:
        """
        Check if the given key and value are present in BOM.

        Args:
            key (str): The key to check.
            value (str): The value to check.
            titles (List[Tuple]): List of tuples representing BOM titles.
            main_title (List[Tuple]): Main title data.

        Returns:
            bool: True if the key and value are present in BOM, False otherwise.
        """
        if key == BomScheduleKey.PART_NUMBER.value:
            return value == main_title[0][3]
        if key == BomScheduleKey.RADAR_PART_NUMBER.value:
            return any(value == item[3] and "radar" in item[0].lower() for item in titles)
        if key == BomScheduleKey.SW_PART_NUMBER.value:
            return any(value == item[3] and "software" in item[0].lower() for item in titles)
        if key == BomScheduleKey.SW_VERSION.value:
            return self._check_sw_version(value, main_title[0][2])
        if key == BomScheduleKey.CONFIG_INI_DATA_SET.value:
            return any(value == item[3] and ("dataset" in item[2].lower() or "data set" in item[2].lower()) for item in titles)
        if key in [BomScheduleKey.PRODUCT_ID_LABEL_PART_NUMBER.value, BomScheduleKey.PART_NUMBER_LABEL.value, BomScheduleKey.LABEL.value, BomScheduleKey.REMAN_LABEL.value]:
            return any(value == item[3] and "label" in item[0].lower() for item in titles)
        if key == BomScheduleKey.JUMPER_HARNESS.value:
            return self._check_jumper_harness(value, titles)
        if key == BomScheduleKey.COVER.value:
            return self._check_cover(value, titles)
        if key == BomScheduleKey.BRACKET.value:
            return any(value == item[3] and "bracket" in item[0].lower() for item in titles)
        if key == BomScheduleKey.CAMERA_PART_NUMBER.value:
            return any(value == item[3] and "camera" in item[0].lower() for item in titles)
        if key == BomScheduleKey.BOOT_SOFTWARE_PART_NUMBER.value:
            return any(value == item[3] and "software" in item[0].lower() and "boot software" in item[2].lower() for item in titles)
        if key == BomScheduleKey.MAIN_SOFTWARE_PART_NUMBER.value:
            return any(value == item[3] and "software" in item[0].lower() and not "boot software" in item[2].lower() for item in titles)
        if key == BomScheduleKey.REMAN_CAMERA_PART_NUMBER.value:
            return any(value == item[3] and "camera" in item[0].lower() for item in titles)
        return False

    def _check_sw_version(self, value: str, main_title_value: str) -> bool:
        """
        Check if the SW version is correctly represented in the BOM.

        Args:
            value (str): The SW version to check.
            main_title_value (str): The main title value from the BOM to compare against.

        Returns:
            bool: True if the SW version is correct, False otherwise.
        """
        pattern = r"^(BX(\d+))(.*)"
        match = re.match(pattern, value)
        if not match:
            return False
        bx_digits, character = match.group(1), match.group(3)
        character = character.upper()
        if character == "A":
            return bx_digits.upper() in main_title_value.upper() and "AUTOBAUD" in main_title_value.upper()
        if character == "H":
            return bx_digits.upper() in main_title_value.upper() and "AUTOBAUD" not in main_title_value.upper() and "500K" in main_title_value.upper()
        if character == "L":
            return bx_digits.upper() in main_title_value.upper() and "AUTOBAUD" not in main_title_value.upper() and "250K" in main_title_value.upper()
        return False

    def _check_jumper_harness(self, value: str, titles: List[Tuple]) -> bool:
        """
        Check if the jumper harness value is correctly represented in the BOM.

        Args:
            value (str): The jumper harness value to check.
            titles (List[Tuple]): List of tuples representing BOM titles.

        Returns:
            bool: True if the jumper harness value is correct, False otherwise.
        """
        if value == "YES":
            return any("harness" in item[0].lower() for item in titles)
        return any(value == item[3] and "harness" in item[0].lower() for item in titles) if re.match(r"^K\d+", value) else False

    def _check_cover(self, value: str, titles: List[Tuple]) -> bool:
        """
        Check if the cover value is correctly represented in the BOM.

        Args:
            value (str): The cover value to check.
            titles (List[Tuple]): List of tuples representing BOM titles.

        Returns:
            bool: True if the cover value is correct, False otherwise.
        """
        if value == "YES":
            return any("cover" in item[0].lower() for item in titles)
        return any(value == item[3] and "cover" in item[0].lower() for item in titles) if re.match(r"^K\d+", value) else False

    def _post_process_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Post-process the DataFrame to handle special cases and formatting.

        Args:
            df (pd.DataFrame): The DataFrame to post-process.

        Returns:
            pd.DataFrame: The post-processed DataFrame.
        """
        df["Comparison vs BOM"] = df.apply(
            lambda row: (
                "--------"
                if row["Schedule Value"].upper() == "NO"
                else row["Comparison vs BOM"]
            ),
            axis=1,
        )
        parameters_to_check = [
            BomScheduleKey.RADAR_PART_NUMBER.value,
            BomScheduleKey.SW_PART_NUMBER.value,
            BomScheduleKey.CONFIG_INI_DATA_SET.value,
            BomScheduleKey.JUMPER_HARNESS.value,
            BomScheduleKey.PRODUCT_ID_LABEL_PART_NUMBER.value,
            BomScheduleKey.PART_NUMBER_LABEL.value,
            BomScheduleKey.BOOT_SOFTWARE_PART_NUMBER.value,
            BomScheduleKey.MAIN_SOFTWARE_PART_NUMBER.value,
            BomScheduleKey.SERVICE_LABEL.value,
            BomScheduleKey.REMAN_CAMERA_PART_NUMBER.value,
            BomScheduleKey.REMAN_LABEL.value,
            BomScheduleKey.LABEL.value,
            BomScheduleKey.BRACKET.value,
            BomScheduleKey.CAMERA_PART_NUMBER.value,
        ]
        df["Comparison vs BOM"] = df.apply(
            lambda row: (
                "--------"
                if row["Parameter"] in parameters_to_check
                and not re.match(r"^K\d+", row["Schedule Value"])
                else row["Comparison vs BOM"]
            ),
            axis=1,
        )
        return df

    def style_schedule_result(self, comparison_result: pd.DataFrame) -> pd.DataFrame.style:
        """
        Apply styling to the comparison result DataFrame.

        Args:
            comparison_result (pd.DataFrame): DataFrame with comparison results.

        Returns:
            pd.DataFrame.style: Styled DataFrame.
        """
        def highlight_mismatched_values(val):
            color = "red" if val["Comparison vs BOM"] == "Incorrect" else "black"
            font_weight = "bold" if val["Comparison vs BOM"] == "Incorrect" else "normal"
            return [f"color: {color}; font-weight: {font_weight}" for _ in val]

        return comparison_result.style.apply(
            highlight_mismatched_values,
            axis=1,
            subset=["Parameter", "Schedule Value", "Comparison vs BOM"]
        ).set_table_styles(
            [
                {"selector": "th", "props": [("text-align", "left"), ("font-weight", "bold")]},
                {"selector": "td:nth-child(2)", "props": [("width", "250px"), ("text-align", "left"), ("word-break", "break-word")]},
                {"selector": "td:nth-child(3)", "props": [("width", "400px"), ("text-align", "left"), ("word-break", "break-word")]},
                {"selector": "td:nth-child(4)", "props": [("width", "400px"), ("text-align", "left"), ("word-break", "break-word")]},
            ]
        )
