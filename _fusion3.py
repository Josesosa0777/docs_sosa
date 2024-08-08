import re
import time
from typing import List, Optional, Tuple, Union, Dict

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
        self.utils = utils
        self.extra_time_sec = int(extra_time)
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
        type_dropdown: DropdownType,
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

        if type_dropdown in {DropdownType.FLR_25_BRACKET, DropdownType.FLC_25_BRACKET}:
            if not self.bracket_in_title:
                raise ValueError(
                    "No specific required element defined for Bracket Assembly. Expected to have 'Bracket Assembly' in the Title ID Value attribute"
                )
            self.element_name = "Bracket"
        elif type_dropdown == DropdownType.FLC_25_COVER:
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

    def is_AM_bracket(self, driver: WebDriver) -> bool:
        """
        Determine if the product is an AM bracket.

        Args:
            driver: The Selenium WebDriver instance used for automation.

        Returns:
            bool: True if the product is an AM bracket, False otherwise.
        """
        sales_chann_limit = self.utils.select_product_info(driver)
        self.is_am_bracket = sales_chann_limit.lower().strip() == "am only"
        return self.is_am_bracket

    def expand_elements_get_titles(self) -> List[Tuple[str, str, str, str]]:
        """
        Expand elements based on conditions and retrieve their titles.

        Returns:
            List[Tuple[str, str, str, str]]: A list of tuples containing element titles.
        """
        time.sleep(0.5 + 0.25 * self.extra_time_sec)
        separator = self.driver.find_element(
            By.XPATH,
            "//*[@id='main-view']/div/div[2]/div[1]/div/div/div/div/div[3]/div/div/div[1]/div/div/main/div[2]",
        )
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
            titles = self._handle_radar_titles(titles)
        if self.camera_in_title:
            titles = self._handle_camera_titles(titles)
        if self.bracket_in_title:
            titles = self._handle_bracket_titles(titles)

        actions.click_and_hold(separator).move_by_offset(-move_x, 0).release().perform()
        return titles

    def _handle_radar_titles(self, titles: List[Tuple[str, str, str, str]]) -> List[Tuple[str, str, str, str]]:
        if self.letter_identifier == LetterIdentifier.R:
            titles = self._expand_bracket_elements(titles, "radar")
        if self.letter_identifier == LetterIdentifier.SC:
            titles = self._expand_child_elements(titles, "radar")
        return titles

    def _handle_camera_titles(self, titles: List[Tuple[str, str, str, str]]) -> List[Tuple[str, str, str, str]]:
        if self.letter_identifier == LetterIdentifier.R:
            titles = self._expand_bracket_elements(titles, "camera")
        if self.letter_identifier == LetterIdentifier.SC:
            titles = self._expand_child_elements(titles, "camera")
        return titles

    def _handle_bracket_titles(self, titles: List[Tuple[str, str, str, str]]) -> List[Tuple[str, str, str, str]]:
        if self.type_dropdown == DropdownType.FLR_25_BRACKET:
            titles = self._expand_bracket_elements(titles, "bracket")
        return titles

    def _expand_bracket_elements(self, titles: List[Tuple[str, str, str, str]], element_name: str) -> List[Tuple[str, str, str, str]]:
        idx = next(
            (
                index
                for index, (title, _, _, _) in enumerate(titles)
                if element_name in title.lower() and "bracket" in title.lower()
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
        return titles

    def _expand_child_elements(self, titles: List[Tuple[str, str, str, str]], element_name: str) -> List[Tuple[str, str, str, str]]:
        idx = next(
            (
                index
                for index, (title, _, _, _) in enumerate(titles)
                if element_name in title.lower()
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
        return titles

    def get_required_elements(self, element_name: str) -> List[str]:
        """
        Get required elements based on element name and letter identifier.

        Args:
            element_name: The name of the element (e.g., Radar, Camera, Bracket, Cover).

        Returns:
            List[str]: A list of required elements.
        """
        elements = []
        if element_name == "Radar":
            elements = self._get_radar_elements()
        elif element_name == "Camera":
            elements = self._get_camera_elements()
        elif element_name == "Bracket":
            elements = self._get_bracket_elements()
        elif element_name == "Cover":
            elements = ["Cover"]
        return elements

    def _get_radar_elements(self) -> List[str]:
        if self.letter_identifier == LetterIdentifier.X:
            return ["Dataset", "Software", "Boot Software"] + (
                ["with CAN termination"]
                if self.chbx_resistor
                else ["wo CAN termination"]
            )
        if self.letter_identifier == LetterIdentifier.R:
            return (
                ["Dataset", "Software", "Boot Software"]
                + (["Screw", "Bracket", "Nut"] if self.chbx_bracket else [])
                + (
                    ["with CAN termination"]
                    if self.chbx_resistor
                    else ["wo CAN termination"]
                )
            )
        if self.letter_identifier == LetterIdentifier.SC:
            return ["Dataset", "Software", "Boot Software"] + (
                ["with CAN termination"]
                if self.chbx_resistor
                else ["wo CAN termination"]
            )

    def _get_camera_elements(self) -> List[str]:
        if self.letter_identifier in {LetterIdentifier.X, LetterIdentifier.SC}:
            return ["Dataset", "Software", "Boot Software", "Camera"]
        if self.letter_identifier == LetterIdentifier.R:
            return ["Dataset", "Software", "Boot Software", "Camera"]

    def _get_bracket_elements(self) -> List[str]:
        if self.type_dropdown == DropdownType.FLR_25_BRACKET:
            return (
                ["Screw", "Bracket", "Nut"]
                if self.is_am_bracket
                else ["Bracket", "Nut"]
            )
        if self.type_dropdown == DropdownType.FLC_25_BRACKET:
            return ["Bracket", "Bracket Assembly", "Worm", "Tape"]

    def get_bom_in_content(self) -> List[Dict[str, Union[str, pd.DataFrame]]]:
        """
        Get titles in Content tab and generate a BOM report.

        Returns:
            List[Dict[str, Union[str, pd.DataFrame]]]: A list containing the BOM report and DataFrame.
        """
        titles = self.expand_elements_get_titles()
        time.sleep(1 + 0.25 * self.extra_time_sec)

        required_elements = self.get_required_elements(self.element_name)
        missing_elements = [
            element
            for element in required_elements
            if not any(element in title for title, _, _, _ in titles)
        ]

        data = {
            "ID": [],
            "Title": [],
            "Description": [],
            "Quantity": [],
            "Expected_value": [],
        }

        self._populate_bom_data(data, titles, required_elements, missing_elements)
        df = pd.DataFrame(data)
        styled_df = self._apply_dataframe_styles(df)

        msg_pdf = self._generate_bom_report(missing_elements, styled_df)
        return [{"msg": msg_pdf, "df": df}]

    def _populate_bom_data(self, data: Dict[str, List[Union[str, int]]], titles: List[Tuple[str, str, str, str]], required_elements: List[str], missing_elements: List[str]) -> None:
        for element in required_elements:
            expected_value = 3 if element in {"Nut", "Screw"} else 1
            for title, quantity, description, _id in titles:
                qty = int(float(quantity))
                if self.type_dropdown == DropdownType.FLC_25_BRACKET and element == "Bracket":
                    if element.lower() in title.lower() and "assembly" not in title.lower():
                        self._append_bom_data(data, _id, title, description, qty, expected_value)
                else:
                    if element in title or (element in {"wo CAN termination", "with CAN termination"} and element.lower() in description.lower()):
                        self._append_bom_data(data, _id, title, description, qty, expected_value)
        self._add_can_termination_elements(data, titles, required_elements)
        self._add_screw_elements_for_sc(data, titles)

    def _append_bom_data(self, data: Dict[str, List[Union[str, int]]], _id: str, title: str, description: str, qty: int, expected_value: int) -> None:
        data["ID"].append(_id)
        data["Title"].append(title)
        data["Description"].append(description)
        data["Quantity"].append(qty)
        data["Expected_value"].append(expected_value)

    def _add_can_termination_elements(self, data: Dict[str, List[Union[str, int]]], titles: List[Tuple[str, str, str, str]], required_elements: List[str]) -> None:
        for title, quantity, description, _id in titles:
            if ("wo CAN termination" in required_elements and "with CAN termination".lower() in description.lower()) or ("with CAN termination" in required_elements and "wo CAN termination".lower() in description.lower()):
                self._append_bom_data(data, _id, title, description, int(float(quantity)), 1)

    def _add_screw_elements_for_sc(self, data: Dict[str, List[Union[str, int]]], titles: List[Tuple[str, str, str, str]]) -> None:
        if self.letter_identifier == LetterIdentifier.SC:
            for title, quantity, description, _id in titles:
                if "screw" in title.lower():
                    self._append_bom_data(data, _id, title, description, int(float(quantity)), 3)
                    break

    def _apply_dataframe_styles(self, df: pd.DataFrame) -> pd.Styler:
        def highlight_mismatched_values(val):
            color = "red" if val["Quantity"] != val["Expected_value"] else "black"
            font_weight = "bold" if val["Quantity"] != val["Expected_value"] else "normal"
            return [f"color: {color}; font-weight: {font_weight}" for _ in val]

        return df.style.apply(
            highlight_mismatched_values,
            axis=1,
            subset=["Title", "Description", "Quantity", "Expected_value"],
        ).set_table_styles(
            [
                {"selector": "td:nth-child(2)", "props": [("width", "140px")]},
                {"selector": "td:nth-child(3)", "props": [("width", "180px")]},
                {"selector": "td:nth-child(4)", "props": [("width", "450px")]},
                {"selector": "td:nth-child(5)", "props": [("width", "120px")]},
                {"selector": "td:nth-child(6)", "props": [("width", "120px")]},
            ]
        )

    def _generate_bom_report(self, missing_elements: List[str], styled_df: pd.Styler) -> str:
        msg_pdf = ""
        if missing_elements:
            msg_missing = f"Missing elements: {missing_elements} for the part number {self.part_number}\n"
            print(f"\033[1m\033[48;5;208m{msg_missing}\033[0m")
            msg_pdf += msg_missing
            display(styled_df)
        else:
            required_elements = self.get_required_elements(self.element_name)
            msg_success_init = "\nSUCCESSFUL:"
            msg_success = f"\nThe required elements: {required_elements}, are present for the part {self.part_number}\n"
            print(f"\033[1;30;48;5;34m{msg_success_init}\033[0m\033[1m{msg_success}\033[0m")
            msg_pdf += msg_success_init + msg_success
            display(styled_df)
        return msg_pdf

    def validate_required_docs(self, documents: dict) -> List[Dict[str, Union[str, pd.DataFrame]]]:
        """
        Validate required documents for a part number.

        Args:
            documents (dict): A dictionary containing document details.

        Returns:
            List[Dict[str, Union[str, pd.DataFrame]]]: Validation results.
        """
        msg_pdf = ""
        if not documents:
            msg_pdf = f"No documents were found in {self.part_number}"
            return [{"msg": msg_pdf, "df": pd.DataFrame()}]

        documents_df = pd.DataFrame.from_dict(documents, orient="index")
        required_docs, optional_docs, required_states = self._get_required_docs_and_states()
        required_documents = documents_df.query(
            "document_kind in @required_docs and lifecycle_state in @required_states"
        )

        present_docs_list, missing_docs = self._check_required_documents(required_documents, required_docs)
        service_data_missing_langs = self._check_service_data_languages(required_docs, required_documents, missing_docs)

        self._generate_missing_docs_report(missing_docs, required_docs, service_data_missing_langs, msg_pdf)
        return self._generate_present_docs_report(present_docs_list, msg_pdf)

    def _get_required_docs_and_states(self) -> Tuple[List[str], List[str], List[str]]:
        required_states = ["Working", "Proposed", "Rejected", "Approved", "Released"]
        if self.element_name == "Radar":
            if self.letter_identifier == LetterIdentifier.R:
                required_docs = ["Assembly Drawing", "Installation Drawing", "Component Drawing", "Service Data"]
            else:
                required_docs = ["Assembly Drawing", "Component Drawing", "Service Data"]
        elif self.element_name == "Bracket":
            required_docs, optional_docs = self._get_bracket_docs()
        elif self.element_name == "Cover":
            required_docs = ["Component Drawing"] if self.is_am_bracket else ["Installation Drawing", "Component Drawing"]
            optional_docs = []
        else:  # Case Camera:
            required_docs = ["Assembly Drawing", "Installation Drawing", "Service Data"] if self.letter_identifier == LetterIdentifier.R else ["Assembly Drawing", "Service Data"]
            optional_docs = []

        return required_docs, optional_docs, required_states

    def _get_bracket_docs(self) -> Tuple[List[str], List[str]]:
        if self.type_dropdown == DropdownType.FLR_25_BRACKET:
            optional_docs = ["Feasibility Agreement (FeA) (DMI)", "Mechanical TR Document", "Electrical TR Document", "Test Report"]
            required_docs = ["Material Specification", "Supplier PPAP", "Order Drawing", "DVP"] + optional_docs
            if self.is_am_bracket:
                required_docs = []
                no_doc_msg = "\nNo documents are required for this part."
                print(f"\033[1m{no_doc_msg}\033[0m")
        if self.type_dropdown == DropdownType.FLC_25_BRACKET:
            optional_docs = ["Electrical TR Document", "Process Data Sheet"]
            required_docs = ["Technical Customer Document", "Assembly Drawing", "Installation Drawing"] + optional_docs
        return required_docs, optional_docs

    def _check_required_documents(self, required_documents: pd.DataFrame, required_docs: List[str]) -> Tuple[List[Dict[str, Union[str, str]]], List[str]]:
        present_docs_list = []
        missing_docs = []
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
            if not doc_exists:
                missing_docs.append(doc_type)
        return present_docs_list, missing_docs

    def _check_service_data_languages(self, required_docs: List[str], required_documents: pd.DataFrame, missing_docs: List[str]) -> Dict[str, bool]:
        service_data_missing_langs = {"US": True}
        if "Service Data" in required_docs:
            for index, row in required_documents.iterrows():
                if row["document_kind"] == "Service Data":
                    title_langs = [lang.strip() for lang in index.split(",")[1:]]
                    for lang in service_data_missing_langs:
                        if lang in title_langs:
                            service_data_missing_langs[lang] = False
                        if "EN" in title_langs:
                            service_data_missing_langs["US"] = False
        return service_data_missing_langs

    def _generate_missing_docs_report(self, missing_docs: List[str], required_docs: List[str], service_data_missing_langs: Dict[str, bool], msg_pdf: str) -> None:
        if missing_docs:
            msg_missing = f"\nThe following required documents are missing in the part {self.part_number}:\n"
            msg_pdf += msg_missing
            print(f"\033[1m\033[48;5;208m{msg_missing}\033[0m")
            for doc in missing_docs:
                doc_msg = f"- {doc}"
                print(f"\033[1;38;2;255;0;0m{doc_msg}\033[0m")
                msg_pdf += f"{doc_msg}\n"
            if any("Service Data" in doc for doc in missing_docs):
                serv_data_msg = "\nPlease validate if the 'Service Data' document is required.\n"
                print(f"\033[1m\033[38;2;255;165;0m{serv_data_msg}\033[0m")
                msg_pdf += serv_data_msg
            for element in optional_docs:
                if any(element in doc for doc in missing_docs):
                    any_element_msg = f"\nPlease validate if the '{element}' document is required."
                    print(f"\033[1m\033[38;2;255;165;0m{any_element_msg}\033[0m")
                    msg_pdf += any_element_msg

    def _generate_present_docs_report(self, present_docs_list: List[Dict[str, Union[str, str]]], msg_pdf: str) -> List[Dict[str, Union[str, pd.DataFrame]]]:
        if present_docs_list:
            present_doc_types = pd.DataFrame(present_docs_list).sort_values(by="Lifecycle_State")
            msg_present_docs = f"\n\nPresent documents in {self.part_number}:\n"
            print(f"\033[1m{msg_present_docs}\033[0m")
            msg_pdf += msg_present_docs
            df_aligned = present_doc_types.style.set_table_styles([dict(selector="th", props=[("text-align", "center")])])
            df_aligned.set_properties(**{"text-align": "left"}).hide(axis="index")
            display(df_aligned)
            return [{"msg": msg_pdf, "df": present_doc_types}]
        return [{"msg": msg_pdf, "df": pd.DataFrame()}]

    def display_dataframe_width(self, df: pd.DataFrame, title: str, column_widths: Dict[str, int]) -> None:
        """
        Display DataFrame with specified column widths.

        Args:
            df: The DataFrame to display.
            title: The title to display above the DataFrame.
            column_widths: A dictionary specifying the column widths.
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

    def validate_aftermarket_product(self, progress_bar) -> List[Dict[str, Union[str, pd.DataFrame]]]:
        """
        Validate aftermarket product documents and display them.

        Args:
            progress_bar: A progress bar widget for displaying progress.

        Returns:
            List[Dict[str, Union[str, pd.DataFrame]]]: Validation results.
        """
        if self.type_dropdown in {DropdownType.FLC_25_BRACKET, DropdownType.FLR_25_BRACKET, DropdownType.FLC_25_COVER}:
            docs_aftermarket = self.utils.documents_aftermarket_product_for(self.driver) if self.is_am_bracket else None
            docs_aftermarket = self.utils.documents_aftermarket_product(self.driver) if not docs_aftermarket else docs_aftermarket
        else:
            docs_aftermarket = self.utils.documents_aftermarket_product(self.driver)

        progress_bar.update(15)
        required_doc_aftermarket = [self.element_name]
        lifecycle_state = ["Working", "Proposed", "Rejected", "Approved", "Released"]

        msg_pdf_am, aftermarket_df, missing_aftermarket_elements = self._process_aftermarket_docs(docs_aftermarket, required_doc_aftermarket, lifecycle_state)
        if missing_aftermarket_elements:
            self._generate_missing_aftermarket_report(missing_aftermarket_elements, msg_pdf_am)
        
        column_widths = {"Document": 180, "Title_ID_value": 200, "Lifecycle_State": 100}
        if not aftermarket_df.empty:
            msg_pdf_am += f"Results for AfterMarket Product in part {self.part_number}:\n"
            self.display_dataframe_width(aftermarket_df, "**AfterMarket Product:**", column_widths)
        
        return [{"msg": msg_pdf_am, "df": aftermarket_df}]

    def _process_aftermarket_docs(self, docs_aftermarket: Optional[Dict[str, dict]], required_doc_aftermarket: List[str], lifecycle_state: List[str]) -> Tuple[str, pd.DataFrame, List[str]]:
        msg_pdf_am = ""
        if docs_aftermarket:
            aftermarket_df = pd.DataFrame.from_dict(docs_aftermarket, orient="index")
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
            aftermarket_df = pd.DataFrame(columns=["Document", "Title_ID_value", "Lifecycle_State"])
            missing_aftermarket_elements = required_doc_aftermarket
        return msg_pdf_am, aftermarket_df, missing_aftermarket_elements

    def _generate_missing_aftermarket_report(self, missing_aftermarket_elements: List[str], msg_pdf_am: str) -> None:
        msg_missing_am = "The following required documents in Aftermarket Product are missing:\n"
        print(f"\033[1m\033[48;5;208m{msg_missing_am}\033[0m")
        msg_pdf_am += msg_missing_am
        for doc in missing_aftermarket_elements:
            doc_msg_am = f"- {doc}"
            print(f"\033[1;38;2;255;0;0m{doc_msg_am}\033[0m")
            msg_pdf_am += f"{doc_msg_am}\n"
        print("\n")
