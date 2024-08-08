import logging
import os
import re
import time
from typing import Optional, Tuple, List, Union, Dict

from utils import Utils
from fusion3 import Fusion3
from legacy import Legacy
from comparison import ComparisonExcel
from reports import PDFReportGenerator

import ipywidgets as widgets
from IPython.display import clear_output, display

from selenium.webdriver.remote.webdriver import WebDriver
from selenium.common.exceptions import TimeoutException
from tqdm.notebook import tqdm
from enums import FileExtension, ValidationType, LetterIdentifier, LifeCycleStatus, TypeOption, ScheduleOption, DropdownType

# Define constants
SELECT_TYPE_OPTION = "-- Select a type --"
SELECT_OPTION = "-- Select an option --"

class CVSplmAutomation:
    """Class for automating tasks on CVSplm website."""

    def __init__(self, 
                 reference_file_camera: str, 
                 reference_file_radar: str, 
                 reference_file_camera_f3: str, 
                 reference_file_radar_f3: str, 
                 extra_time: int = 0):
        """
        Initialize the CVSplmAutomation class.

        Args:
            reference_file_camera: Path to the reference file for camera.
            reference_file_radar: Path to the reference file for radar.
            reference_file_camera_f3: Path to the reference file for camera (F3).
            reference_file_radar_f3: Path to the reference file for radar (F3).
            extra_time: Optional additional time (in seconds) to adjust delays for actions.
        """
        self.driver: Optional[WebDriver] = None
        self.reference_file_camera = reference_file_camera
        self.reference_file_radar = reference_file_radar
        self.reference_file_camera_f3 = reference_file_camera_f3
        self.reference_file_radar_f3 = reference_file_radar_f3
        self.part_number: Optional[str] = None
        self.extra_time_sec: int = extra_time
        (
            self.part_number_input,
            self.type_dropdown,
            self.option_dropdown,
            self.chbx_bracket,
            self.chbx_resistor,
            self.chbx_schedule,
            self.schedule_dropdown,
        ) = self.setup_widgets()
        self.title_id: Optional[str] = None
        self.title_id_value: Optional[str] = None
        self.life_cycle_state: Optional[LifeCycleStatus] = None
        self.letter_identifier: Optional[LetterIdentifier] = None
        self.main_div_path = "//main/div[1]/div/div[2]/div/aw-splm-table/div/div[2]/div[3]/div[2]/div/div"
        self.child_div_path = "//*[@id='occTreeTable']/div[2]/div[2]/div[2]/div//div"
        self.utils = Utils(extra_time)
        self.fusion3 = Fusion3(self.utils, extra_time)
        self.legacy = Legacy(self.utils, extra_time)
        self.ini_comparison = ComparisonExcel()
        self.report_generator = PDFReportGenerator()
        self.results_to_pdf = []
        self.validation_type: Optional[ValidationType] = None
        self.total_steps = 100
        self.url = "https://cvsplm.corp.knorr-bremse.com:14432/#/showHome"
        self.btn_generate_pdf = widgets.Button(
            description="Generate PDF Report",
            disabled=True,
            layout=widgets.Layout(width="150px", margin="0 0 0 20px"),
        )
        self.btn_generate_pdf.on_click(self.on_generate_pdf_click)
        self.output = widgets.Output()

    def on_generate_pdf_click(self, b: widgets.Button) -> None:
        """
        Handle the click event for the PDF generation button.

        Args:
            b (widgets.Button): The button instance that triggered the click event.
        """
        with self.output:
            clear_output(wait=True)
            part_identifier = (
                self.part_number
                if self.validation_type != ValidationType.SCHEDULE
                else self.schedule_dropdown.value
            )
            validation_type_enum = ValidationType[self.validation_type]
            self.report_generator.generate_pdf(
                self.results_to_pdf, validation_type_enum, part_identifier
            )

    def _validate_reference_files(self) -> None:
        """
        Validate the existence of reference files.
        """
        self._validate_file(self.reference_file_camera, FileExtension.XLSX)
        self._validate_file(self.reference_file_radar, FileExtension.XLSX)
        self._validate_file(self.reference_file_camera_f3, FileExtension.INI)
        self._validate_file(self.reference_file_radar_f3, FileExtension.INI)

    def _validate_file(self, file_path: str, expected_extension: FileExtension) -> None:
        """
        Validate if the file exists and has the expected extension.

        Args:
            file_path (str): The path to the file to validate.
            expected_extension (FileExtension): The expected file extension.

        Raises:
            ValueError: If the file does not exist or does not have the expected extension.
        """
        if not os.path.exists(file_path) or not file_path.lower().endswith(expected_extension.value):
            raise ValueError(
                f"\033[48;5;208m\nInvalid reference file: '{file_path}'.\nThe file should exist and be in {expected_extension.value} extension.\n\033[0m"
            )

    def _determine_reference_file(self) -> Optional[str]:
        """
        Determine the appropriate reference file based on the title and dropdown type.

        Returns:
            Optional[str]: The path to the appropriate reference file, or None if no file is selected.
        """
        if not self.title_id_value or not self.type_dropdown.value:
            return None

        title_lower = self.title_id_value.lower()
        type_value = self.type_dropdown.value

        if "radar" in title_lower:
            if type_value == DropdownType.FLR_21.value:
                return self.reference_file_radar
            elif type_value == DropdownType.FLR_25.value:
                return self.reference_file_radar_f3
        elif "camera" in title_lower:
            if type_value == DropdownType.FLC_20.value:
                return self.reference_file_camera
            elif type_value == DropdownType.FLC_25.value:
                return self.reference_file_camera_f3

    def search_part(self) -> bool:
        """
        Search for a part using the provided part number and update class attributes with search results.

        Returns:
            True if the part was found, False otherwise.
        """
        self.utils.search_for_part(self.driver, self.part_number)
        search = self.utils.select_from_search(self.driver, self.part_number)
        if search:
            (
                self.title_id,
                self.title_id_value,
                life_cycle_state,
                self.type_number,
                self.customers,
            ) = self.utils.get_attr_from_part_search(self.driver)
            try:
                self.life_cycle_state = LifeCycleStatus(life_cycle_state)
            except ValueError:
                print(f"Invalid life cycle status received: {life_cycle_state}")
                self.life_cycle_state = None
            
            return True
        return False

    def start_browser_interaction(self, progress_bar: Optional[widgets.FloatProgress], process_name: str) -> bool:
        """
        Initializes interaction with the browser of AW and performs part search.

        Args:
            progress_bar (Optional[widgets.FloatProgress]): Optional progress bar widget for displaying progress.
            process_name (str): The name of the process (e.g., "INI", "BOM", etc.).

        Returns:
            bool: True if the part was found, False otherwise.
        """
        try:
            print("\nInitializing interaction with the browser of AW ...\n")
            steps_progress_bar = {
                "INI": [20, 10],
                "BOM": [25, 25],
                "CONTENT": [25, 25],
                "AM": [30, 25],
                "ATTRIBUTES": [25, 25],
            }
            if process_name not in steps_progress_bar:
                raise ValueError(f"Invalid process name: '{process_name}'")

            self.driver = self.utils.open_url(self.url)
            progress_bar.update(steps_progress_bar[process_name][0])
            self.utils.search_for_part(self.driver, self.part_number)
            search_success = self.utils.select_from_search(self.driver, self.part_number)
            progress_bar.update(steps_progress_bar[process_name][1])
            if search_success:
                self.utils.expand_sections(self.driver)
                (
                    self.title_id,
                    self.title_id_value,
                    self.life_cycle_state,
                    self.type_number,
                    self.customers,
                ) = self.utils.get_attr_from_part_search(self.driver)
                return True
            return False
        except TimeoutException:
            print(
                f"\033[1m\033[48;5;208m\nA TimeoutException was raised. Please validate access to Active Workspace and try again.\n\033[0m"
            )
            return False
        except Exception as e:
            print("\nAn error occurred during browser interaction initialization.\n")
            raise e

    def execute_sc_content(self, progress_bar: widgets.FloatProgress) -> List[str]:
        """
        Execute the SC Content process.

        Args:
            progress_bar (widgets.FloatProgress): Progress bar widget for displaying progress.

        Returns:
            List[str]: Results of the SC Content execution.
        """
        results = []
        try:
            self.utils.select_content(self.driver)
            partX = self.utils.get_partX_name_in_content(self.driver)
            self.utils.select_info(self.driver)
            time.sleep(2 + self.extra_time_sec)
            documents = self.utils.find_docs(self.driver, self.part_number)
            progress_bar.update(20)
            results.extend(self.initialize_validate_content(documents))
            progress_bar.update(5)
            if partX:
                self.part_number = partX
                search = self.search_part()
                if search:
                    time.sleep(2 + self.extra_time_sec)
                    documents = self.utils.find_docs(self.driver, self.part_number)
                    progress_bar.update(20)
                    results.extend(self.initialize_validate_content(documents))
                    progress_bar.update(5)
                else:
                    print(f"\nIt was not possible to find the {self.part_number}")
                    progress_bar.update(25)
            else:
                progress_bar.update(25)
                print(f"\nNo Part X was found in the {self.part_number}")
        except Exception as e:
            logging.error(f"An unexpected error occurred during SC Content execution: {e}")
        
        return results

    def initialize_validate_content(self, documents: List[str]) -> List[str]:
        """
        Initialize and validate content based on the documents provided.

        Args:
            documents (List[str]): List of document identifiers.

        Returns:
            List[str]: Validation results.
        """
        type_value = self.type_dropdown.value

        if type_value in [
            DropdownType.FLR_25.value,
            DropdownType.FLC_25.value,
            DropdownType.FLR_25_BRACKET.value,
            DropdownType.FLC_25_BRACKET.value,
            DropdownType.FLC_25_COVER.value
        ]:
            self.fusion3.initialize_main_params(
                self.part_number,
                self.title_id_value,
                self.driver,
                self.child_div_path,
                self.letter_identifier,
                self.chbx_bracket,
                self.chbx_resistor,
                self.type_dropdown,
            )
            result = self.fusion3.validate_required_docs(documents)
        else:
            self.legacy.initialize_main_params(
                self.part_number,
                self.title_id_value,
                self.driver,
                self.child_div_path,
                self.letter_identifier,
                self.chbx_bracket,
                self.chbx_resistor,
                self.type_dropdown,
            )
            result = self.legacy.validate_required_docs(documents)
        return result

    def start_schedule_interaction(self, progress_bar: widgets.FloatProgress, process_name: str) -> bool:
        """
        Start interaction with the browser for schedule-related tasks.

        Args:
            progress_bar (FloatProgress): Progress bar widget for displaying progress.
            process_name (str): Name of the process to update progress bar.

        Returns:
            bool: True if schedule selection was successful, False otherwise.
        """
        try:
            print("\nInitializing interaction with the browser of AW ...\n")
            steps_progress_bar = {"SCHEDULE": [10, 10]}
            self.driver = self.utils.open_url(self.url)
            progress_bar.update(steps_progress_bar[process_name][0])
            schedule_search_input = self.schedule_dropdown.value.split(",")[0]
            self.utils.search_schedule(self.driver, schedule_search_input)
            search = self.utils.select_schedule_from_search(
                self.driver, self.schedule_dropdown.value
            )
            progress_bar.update(steps_progress_bar[process_name][1])
            if search:
                return True
            return False
        except TimeoutException:
            print(
                "\033[1m\033[48;5;208m\nA TimeoutException was raised. Please ensure you have access to Active Workspace and try again.\n\033[0m"
            )
        except Exception as e:
            print(
                "\033[1m\nAn error occurred in browser interaction initialization\n\033[0m"
            )
            raise e

    # --- Execution Methods ---
    def execute_schedule(self) -> None:
        """
        Execute the schedule process.
        """
        with tqdm(
            total=self.total_steps,
            desc="Progress",
            bar_format="{l_bar}{bar}|   Time: {elapsed}s",
        ) as progress_bar:
            try:
                self.validation_type = ValidationType.SCHEDULE
                search = self.start_schedule_interaction(
                    progress_bar, process_name=self.validation_type.value
                )
                if search:
                    self.process_schedule(progress_bar)
                progress_bar.update(80)
                    
            except Exception as e:
                logging.error(f"An error occurred in schedule execution: {e}")
                print("\nAn error occurred in the execution\n")
                raise e
            finally:
                progress_bar.close()
                self.utils.close_session(self.driver)
        
    def process_schedule(self, progress_bar: widgets.FloatProgress) -> None:
        """
        Process the schedule once interaction is successful.

        Args:
            progress_bar ([widgets.FloatProgress]): Progress bar widget.
        """
        self.legacy.init_schedule(self.driver, self.schedule_dropdown, self.chbx_schedule)
        excel_file_path = self.legacy.download_excel_file()
        progress_bar.update(10)

        if excel_file_path:
            rows_dict = self.ini_comparison.extract_info_schedule(excel_file_path)
            if rows_dict:
                self.compare_bom_schedule(rows_dict, progress_bar)
            else:
                print("The excel file was not processed correctly.")
                progress_bar.update(70)
        else:
            print("The excel file was not found.")
            progress_bar.update(70)

    def compare_bom_schedule(self, rows_dict: Dict, progress_bar: widgets.FloatProgress) -> None:
        """
        Compare BOM schedule based on extracted information.

        Args:
            rows_dict (Dict): Dictionary containing row information.
            progress_bar ([widgets.FloatProgress]): Progress bar widget.
        """
        parts = ", ".join([part[0] for part in rows_dict["row_info"]])
        print(f"\033[1m\nStarting Comparison in the BOM for the Parts: '{parts}'\n\033[0m")
        
        steps_percentage = 70 / len(rows_dict["row_info"])
        self.results_to_pdf = []

        for row in rows_dict["row_info"]:
            element_data = self.ini_comparison.get_info_red_row(row, rows_dict["headers"])
            if element_data and element_data["PART NUMBER"]:
                self.part_number = element_data["PART NUMBER"]
                self.utils.search_for_part(self.driver, self.part_number)
                search = self.utils.select_from_search(self.driver, self.part_number)

                if search:
                    progress_bar.update(steps_percentage / 3)
                    self.update_part_attributes(progress_bar, steps_percentage)
                    self.results_to_pdf.append(
                        self.legacy.get_bom_schedule_in_content(element_data)
                    )
                    progress_bar.update(steps_percentage / 3)
                else:
                    progress_bar.update(steps_percentage)
            else:
                progress_bar.update(steps_percentage)

        self.btn_generate_pdf.disabled = False

    def update_part_attributes(self, progress_bar: widgets.FloatProgress, steps_percentage: float) -> None:
        """
        Update attributes of the part and select content.

        Args:
            progress_bar ([widgets.FloatProgress]): Progress bar widget.
            steps_percentage (float): Percentage of completion for steps.
        """
        (
            self.title_id,
            self.title_id_value,
            self.life_cycle_state,
            self.type_number,
            self.customers,
        ) = self.utils.get_attr_from_part_search(self.driver)
        self.utils.select_content(self.driver)
        progress_bar.update(steps_percentage / 3)
        self.legacy.initialize_main_params(
            self.part_number,
            self.title_id_value,
            self.driver,
            self.child_div_path,
            self.letter_identifier,
            self.chbx_bracket,
            self.chbx_resistor,
            self.type_dropdown,
        )
        
    def execute_ini_automation(self):
        """Execute the automation process."""
        with tqdm(
            total=self.total_steps,
            desc="Progress",
            bar_format="{l_bar}{bar}|   Time: {elapsed}s",
        ) as progress_bar:
            try:
                self.validation_type = ValidationType.INI
                search = self.start_browser_interaction(
                    progress_bar, process_name=self.validation_type.value
                )
                if search:
                    self.process_ini_automation(progress_bar)
                else:
                    progress_bar.update(70)
            except Exception as e:
                self.validation_type = None
                print("\nAn error occurred in the execution\n")
                raise e
            finally:
                progress_bar.close()
                self.utils.close_session(self.driver)

    def process_ini_automation(self, progress_bar: tqdm) -> None:
        """
        Process the INI automation once browser interaction is successful.

        Args:
            progress_bar (tqdm): Progress bar widget.
        """
        if not self.utils.validate_type_number(self.part_number, self.type_number, self.type_dropdown):
            return

        reference_file = self._determine_reference_file()
        if not reference_file:
            print(f"\033[48;5;208m\nUnable to determine if Part {self.part_number} is 'Camera' or 'Radar'\n\033[0m")
            progress_bar.update(70)
            return

        self.legacy.initialize_main_params(
            self.part_number,
            self.title_id_value,
            self.driver,
            self.child_div_path,
            self.letter_identifier,
            self.chbx_bracket,
            self.chbx_resistor,
            self.type_dropdown,
        )
        self.utils.select_content(self.driver)
        progress_bar.update(15)

        titles = self.legacy.expand_elements_get_titles()
        dataset_idx = self._find_dataset_index(titles)
        if dataset_idx:
            if self.letter_identifier in [LetterIdentifier.N.value, LetterIdentifier.SC.value]:
                dataset_idx += 1
            progress_bar.update(20)
            self.legacy.select_doc_dataset(dataset_idx)
            progress_bar.update(20)
            file_path = self.legacy.download_excel_file()
            if not file_path:
                print("\033[48;5;208m\nNo ini file was found to continue with validations\n\033[0m")
                progress_bar.update(15)
                return
            print("Interaction with AW finished. Starting comparison of files ...")
            self.compare_ini_files(file_path, reference_file, progress_bar, titles)
        else:
            print("\033[48;5;208m\nNo dataset was found to continue with validations\n\033[0m")
            progress_bar.update(35)

    def _find_dataset_index(self, titles: List[Tuple[str, str]]) -> int:
        """
        Find the dataset index in the titles list.

        Args:
            titles (List[Tuple[str, str]]): List of titles.

        Returns:
            int: The dataset index if found, else None.
        """
        return next(
            (i + 3 for i, title in enumerate(titles) if "dataset" in title[0].lower()),
            None
        )
    
    def compare_ini_files(self, file_path: str, reference_file: str, progress_bar: tqdm, titles: List[Tuple[str, str]]) -> None:
        """
        Compare INI files based on the type dropdown value.

        Args:
            file_path (str): Path to the downloaded INI file.
            reference_file (str): Path to the reference file.
            progress_bar (tqdm): Progress bar widget.
            titles (List[Tuple[str, str]]): List of title of documents.
        """
        if self.type_dropdown.value in [DropdownType.FLC_20.value, DropdownType.FLR_21.value]:
            self.results_to_pdf = self.ini_comparison.comparison(file_path, reference_file, self.part_number)
        elif self.type_dropdown.value in [DropdownType.FLC_25.value, DropdownType.FLR_25.value]:
            self.results_to_pdf = self.ini_comparison.comparison_ini(file_path, reference_file, self.part_number, titles)
        self.btn_generate_pdf.disabled = False
        progress_bar.update(15)

    def execute_bom_automation(self) -> None:
        """Execute the BOM process."""
        with tqdm(
            total=self.total_steps,
            desc="Progress",
            bar_format="{l_bar}{bar}|   Time: {elapsed}s",
        ) as progress_bar:
            try:
                self.validation_type = ValidationType.BOM
                search = self.start_browser_interaction(
                    progress_bar, process_name=self.validation_type.value
                )
                if search:
                    if not self.utils.validate_type_number(
                        self.part_number, self.type_number, self.type_dropdown
                    ):
                        return
                    dropdown_value = self.type_dropdown.value
                    if dropdown_value in [
                        DropdownType.FLR_25.value,
                        DropdownType.FLC_25.value,
                        DropdownType.FLR_25_BRACKET.value,
                        DropdownType.FLC_25_BRACKET.value,
                        DropdownType.FLC_25_COVER.value,
                    ]:
                        if dropdown_value in [
                            DropdownType.FLR_25_BRACKET.value,
                            DropdownType.FLC_25_BRACKET.value,
                            DropdownType.FLC_25_COVER.value,
                        ]:
                            self.letter_identifier = (
                                self.utils.get_identifier_bracket_cover(
                                    self.title_id_value, self.type_dropdown
                                )
                            )
                        self.fusion3.initialize_main_params(
                            self.part_number,
                            self.title_id_value,
                            self.driver,
                            self.child_div_path,
                            self.letter_identifier,
                            self.chbx_bracket,
                            self.chbx_resistor,
                            self.type_dropdown,
                        )
                        if dropdown_value in [
                            DropdownType.FLR_25_BRACKET.value,
                            DropdownType.FLC_25_BRACKET.value,
                            DropdownType.FLC_25_COVER.value,
                        ]:
                            self.fusion3.is_AM_bracket(self.driver)
                    self.utils.select_content(self.driver)
                    progress_bar.update(25)
                    if dropdown_value in [
                        DropdownType.FLR_25.value,
                        DropdownType.FLC_25.value,
                        DropdownType.FLR_25_BRACKET.value,
                        DropdownType.FLC_25_BRACKET.value,
                        DropdownType.FLC_25_COVER.value,
                    ]:
                        self.results_to_pdf = self.fusion3.get_bom_in_content()
                        self.fusion3.is_am_bracket = False
                    else:
                        self.legacy.initialize_main_params(
                            self.part_number,
                            self.title_id_value,
                            self.driver,
                            self.child_div_path,
                            self.letter_identifier,
                            self.chbx_bracket,
                            self.chbx_resistor,
                            self.type_dropdown,
                        )
                        self.results_to_pdf = self.legacy.get_bom_in_content()
                    progress_bar.update(25)
                    self.btn_generate_pdf.disabled = False
            except Exception as e:
                self.validation_type = None
                print("\nAn error occurred in the execution\n")
                raise e
            finally:
                progress_bar.close()
                self.utils.close_session(self.driver)

    def execute_content_automation(self) -> None:
        """Execute the content process."""
        with tqdm(
            total=self.total_steps,
            desc="Progress",
            bar_format="{l_bar}{bar}|   Time: {elapsed}s",
        ) as progress_bar:
            try:
                self.validation_type = ValidationType.CONTENT
                search = self.start_browser_interaction(progress_bar, process_name=self.validation_type.value)
                if search:
                    if not self.utils.validate_type_number(
                        self.part_number, self.type_number, self.type_dropdown
                    ):
                        return
                    dropdown_value = self.type_dropdown.value
                    if dropdown_value in [
                        DropdownType.FLR_25_BRACKET.value,
                        DropdownType.FLC_25_BRACKET.value,
                        DropdownType.FLC_25_COVER.value,
                    ]:
                        self.letter_identifier = (
                            self.utils.get_identifier_bracket_cover(
                                self.title_id_value, self.type_dropdown
                            )
                        )
                        is_am_bracket = self.fusion3.is_AM_bracket(self.driver)
                        if (
                            is_am_bracket
                            and dropdown_value == DropdownType.FLR_25_BRACKET.value
                        ):
                            print(
                                f"\033[1mNo documents are required for this part.\n\033[0m"
                            )
                            progress_bar.update(50)
                            self.fusion3.is_am_bracket = False
                            return
                    if self.letter_identifier == LetterIdentifier.SC:
                        self.results_to_pdf = self.execute_sc_content(progress_bar)
                    else:
                        time.sleep(5 + self.extra_time_sec)
                        documents = self.utils.find_docs(self.driver, self.part_number)
                        progress_bar.update(25)
                        self.results_to_pdf = self.initialize_validate_content(
                            documents
                        )
                        progress_bar.update(25)
                    self.btn_generate_pdf.disabled = False
                else:
                    progress_bar.update(50)
            except Exception as e:
                self.validation_type = None
                print("\nAn error occurred in the execution\n")
                raise e
            finally:
                progress_bar.close()
                self.utils.close_session(self.driver)

    def execute_aftermarket_automation(self):
        """Execute the After Market process."""
        with tqdm(
            total=self.total_steps,
            desc="Progress",
            bar_format="{l_bar}{bar}|   Time: {elapsed}s",
        ) as progress_bar:
            try:
                self.validation_type = ValidationType.AM
                search = self.start_browser_interaction(
                    progress_bar, process_name=self.validation_type.value
                )
                if search:
                    if not self.utils.validate_type_number(
                        self.part_number, self.type_number, self.type_dropdown
                    ):
                        return
                    dropdown_value = self.type_dropdown.value
                    if dropdown_value in [
                        DropdownType.FLR_25.value,
                        DropdownType.FLC_25.value,
                        DropdownType.FLR_25_BRACKET.value,
                        DropdownType.FLC_25_BRACKET.value,
                        DropdownType.FLC_25_COVER.value,
                    ]:
                        if dropdown_value in [
                            DropdownType.FLR_25_BRACKET.value,
                            DropdownType.FLC_25_BRACKET.value,
                            DropdownType.FLC_25_COVER.value,
                        ]:
                            self.letter_identifier = (
                                self.utils.get_identifier_bracket_cover(
                                    self.title_id_value, self.type_dropdown
                                )
                            )
                        self.fusion3.initialize_main_params(
                            self.part_number,
                            self.title_id_value,
                            self.driver,
                            self.child_div_path,
                            self.letter_identifier,
                            self.chbx_bracket,
                            self.chbx_resistor,
                            self.type_dropdown,
                        )
                        if dropdown_value in [
                            DropdownType.FLR_25_BRACKET.value,
                            DropdownType.FLC_25_BRACKET.value,
                            DropdownType.FLC_25_COVER.value,
                        ]:
                            self.fusion3.is_AM_bracket(self.driver)
                        self.utils.select_products(self.driver)
                        progress_bar.update(25)
                        self.results_to_pdf = self.fusion3.validate_aftermarket_product(
                            progress_bar
                        )
                        self.fusion3.is_am_bracket = False
                        progress_bar.update(5)
                    else:
                        self.legacy.initialize_main_params(
                            self.part_number,
                            self.title_id_value,
                            self.driver,
                            self.child_div_path,
                            self.letter_identifier,
                            self.chbx_bracket,
                            self.chbx_resistor,
                            self.type_dropdown,
                        )
                        self.utils.select_products(self.driver)
                        progress_bar.update(25)
                        self.results_to_pdf = self.legacy.validate_aftermarket_product(
                            progress_bar
                        )
                        progress_bar.update(5)
                    self.btn_generate_pdf.disabled = False
                else:
                    progress_bar.update(45)
            except Exception as e:
                progress_bar.close()
                print("\nAn error occurred in the execution\n")
                raise e
            finally:
                progress_bar.close()
                self.utils.close_session(self.driver)

    def execute_attributes_validation(self):
        """Execute the Validation of the attributes."""
        with tqdm(
            total=self.total_steps,
            desc="Progress",
            bar_format="{l_bar}{bar}|   Time: {elapsed}s",
        ) as progress_bar:
            try:
                self.validation_type = ValidationType.ATTRIBUTES
                search = self.start_browser_interaction(
                    progress_bar, process_name=self.validation_type.value
                )
                if search:
                    self.utils.validate_type_number(
                        self.part_number, self.type_number, self.type_dropdown
                    )
                    # self.utils.expand_sections(self.driver)
                    progress_bar.update(25)
                    self.results_to_pdf = self.utils.get_attribute_values(
                        self.driver,
                        self.part_number,
                        self.type_dropdown,
                        self.letter_identifier,
                    )
                    progress_bar.update(25)
                    self.btn_generate_pdf.disabled = False
            except Exception as e:
                print("\nAn error occurred in the execution\n")
                raise e
            finally:
                progress_bar.close()
                self.utils.close_session(self.driver)

    # --- Setup ---
    def setup_widgets(self) -> Tuple[widgets.Text, widgets.Dropdown, widgets.Dropdown, widgets.Checkbox, widgets.Checkbox, widgets.Checkbox, widgets.Dropdown]:
        """
        Set up widgets for user interaction.

        Returns:
            Tuple[widgets.Text, widgets.Dropdown, widgets.Dropdown, widgets.Checkbox, widgets.Checkbox, widgets.Checkbox, widgets.Dropdown]:
                A tuple containing initialized widgets: part number input, type dropdown, option dropdown,
                bracket checkbox, resistor checkbox, schedule checkbox, and schedule dropdown.
        """
        part_number_input = widgets.Text(
            value="",
            placeholder="e.g. K285620N000",
            description="Part Number:",
            disabled=False,
            continuous_update=False,
        )

        type_dropdown = widgets.Dropdown(
            options=[
                SELECT_TYPE_OPTION,
                *[dt.value for dt in DropdownType]
            ],
            description="Type:",
        )

        option_dropdown = widgets.Dropdown(
            options=[SELECT_OPTION], description="Option:"
        )

        chbx_bracket = widgets.Checkbox(
            value=False,
            description="Radar w/mounting bracket",
            disabled=False,
            indent=False,
        )
        chbx_bracket.layout.visibility = "hidden"

        chbx_resistor = widgets.Checkbox(
            value=False,
            description="Radar w/CAN termination",
            disabled=False,
            indent=False,
        )
        chbx_resistor.layout.visibility = "hidden"

        chbx_schedule = widgets.Checkbox(
            value=False,
            description="Schedule validations",
            disabled=False,
            indent=False,
        )
        chbx_schedule.layout.margin = "0 0 0 25px"  # margin to the left

        schedule_dropdown = widgets.Dropdown(
            options=[SELECT_OPTION],
            description="Schedule Number:",
            style={"description_width": "180px"},
            layout=widgets.Layout(
                width="550px", visibility="hidden", margin="0 0 0 -200px"
            ),
        )

        # Link widget callbacks
        part_number_input.observe(self.update_part, names="value")
        type_dropdown.observe(self.update_type, names="value")
        option_dropdown.observe(self.update_ecu_options, names="value")
        chbx_schedule.observe(self.select_schedule, names="value")

        return (
            part_number_input,
            type_dropdown,
            option_dropdown,
            chbx_bracket,
            chbx_resistor,
            chbx_schedule,
            schedule_dropdown,
        )

    def update_part(self, change: dict) -> None:
        """Update options based on the part number input."""
        pattern = r"\d+SC\d+"
        if self.type_dropdown.value in [DropdownType.FLC_25.value, DropdownType.FLR_25.value]:
            if re.search(pattern, change["new"]):
                self.option_dropdown.options = [
                    SELECT_OPTION,
                    TypeOption.COMPARE_INI_FILE.value,
                    TypeOption.BOM_STRUCTURE.value,
                    TypeOption.DOCUMENT_VALIDATION.value,
                    TypeOption.ATTRIBUTES_VALIDATION.value,
                ]
            else:
                self.option_dropdown.options = [
                    SELECT_OPTION,
                    TypeOption.COMPARE_INI_FILE.value,
                    TypeOption.BOM_STRUCTURE.value,
                    TypeOption.DOCUMENT_VALIDATION.value,
                    TypeOption.AFTERMARKET_PARTS.value,
                    TypeOption.ATTRIBUTES_VALIDATION.value,
                ]

    def update_type(self, change: dict) -> None:
        """
        Callback function to update options based on type selection.

        Args:
            change (dict): Dictionary containing the new value of the dropdown.
        """
        selected_type = change["new"]
        if selected_type == DropdownType.FLR_25.value:
            self.show_bracket_resistor()
            self.reset_checkboxes(False)
            options = self.get_options_based_on_part_number()
        elif selected_type == DropdownType.FLC_25.value:
            self.hide_bracket_resistor()
            self.reset_checkboxes(False)
            options = self.get_options_based_on_part_number()
        elif selected_type in {DropdownType.FLC_20.value, DropdownType.FLR_21.value}:
            self.hide_bracket_resistor()
            self.reset_checkboxes(False)
            options = [
                SELECT_OPTION,
                TypeOption.COMPARE_INI_FILE.value,
                TypeOption.BOM_STRUCTURE.value,
                TypeOption.DOCUMENT_VALIDATION.value,
                TypeOption.AFTERMARKET_PARTS.value,
                TypeOption.ATTRIBUTES_VALIDATION.value,
            ]
        elif selected_type in {DropdownType.FLR_25_BRACKET.value, DropdownType.FLC_25_BRACKET.value}:
            self.hide_bracket_resistor()
            self.reset_checkboxes(False)
            options = [
                SELECT_OPTION,
                TypeOption.BOM_STRUCTURE.value,
                TypeOption.DOCUMENT_VALIDATION.value,
                TypeOption.AFTERMARKET_PARTS.value,
            ]
        elif selected_type == DropdownType.FLC_25_COVER.value:
            self.hide_bracket_resistor()
            self.reset_checkboxes(False)
            options = [SELECT_OPTION, TypeOption.DOCUMENT_VALIDATION.value]
        else:
            options = [SELECT_OPTION]

        self.option_dropdown.options = options

    def show_bracket_resistor(self) -> None:
        """Show bracket and resistor checkboxes."""
        self.chbx_bracket.layout.visibility = "visible"
        self.chbx_resistor.layout.visibility = "visible"
        self.chbx_resistor.layout.margin = "0 0 0 25px"
        self.chbx_bracket.layout.margin = "0 0 0 -100px"

    def hide_bracket_resistor(self) -> None:
        """Hide bracket and resistor checkboxes."""
        self.chbx_bracket.layout.visibility = "hidden"
        self.chbx_resistor.layout.visibility = "hidden"

    def reset_checkboxes(self, flag: bool) -> None:
        """
        Reset the state of checkboxes.

        Args:
            disable_bracket (bool): Whether to disable the bracket checkbox.
        """
        self.chbx_bracket.value = flag
        self.chbx_resistor.value = flag
        self.chbx_bracket.disabled = flag

    def get_options_based_on_part_number(self) -> List[str]:
        """
        Get options based on the part number.

        Returns:
            List[str]: List of options for the dropdown.
        """
        pattern = r"\d+SC\d+"
        if re.search(pattern, self.part_number_input.value):
            return [
                SELECT_OPTION,
                TypeOption.COMPARE_INI_FILE.value,
                TypeOption.BOM_STRUCTURE.value,
                TypeOption.DOCUMENT_VALIDATION.value,
                TypeOption.ATTRIBUTES_VALIDATION.value,
            ]
        else:
            return [
                SELECT_OPTION,
                TypeOption.COMPARE_INI_FILE.value,
                TypeOption.BOM_STRUCTURE.value,
                TypeOption.DOCUMENT_VALIDATION.value,
                TypeOption.AFTERMARKET_PARTS.value,
                TypeOption.ATTRIBUTES_VALIDATION.value,
            ]

    def update_option_dropdown(self, *options: TypeOption) -> None:
        """
        Update the option dropdown based on provided options.

        Args:
            options (TypeOption): Enums representing the options to be added to the dropdown.
        """
        self.option_dropdown.options = [SELECT_OPTION] + [opt.value for opt in options]

    def update_option_dropdown_based_on_part_number(self) -> None:
        """
        Update options based on the part number and type selection.
        """
        pattern = r"\d+SC\d+"
        if re.search(pattern, self.part_number_input.value):
            self.update_option_dropdown(
                TypeOption.COMPARE_INI_FILE,
                TypeOption.BOM_STRUCTURE,
                TypeOption.DOCUMENT_VALIDATION,
                TypeOption.ATTRIBUTES_VALIDATION
            )
        else:
            self.update_option_dropdown(
                TypeOption.COMPARE_INI_FILE,
                TypeOption.BOM_STRUCTURE,
                TypeOption.DOCUMENT_VALIDATION,
                TypeOption.AFTERMARKET_PARTS,
                TypeOption.ATTRIBUTES_VALIDATION
            )

    def update_ecu_options(self, change: dict) -> None:
        """
        Callback function to update ECU options based on selection.

        Args:
            change (dict): Dictionary containing the new value of the dropdown.
        """
        if change["new"] == "BOM Structure" and self.type_dropdown.value == "FLR-25":
            self.update_bracket_resistor_visibility(True)
        else:
            self.update_bracket_resistor_visibility(False)

    def update_bracket_resistor_visibility(self, flag: bool) -> None:
        """
        Update visibility of bracket and resistor checkboxes.

        Args:
            visible (bool): Determines if checkboxes should be visible or hidden.
        """
        self.chbx_bracket.layout.visibility = "visible" if flag else "hidden"
        self.chbx_resistor.layout.visibility = "visible" if flag else "hidden"
        self.chbx_resistor.layout.margin = "0 0 0 25px"
        self.chbx_bracket.layout.margin = "0 0 0 -100px"

    def select_schedule(self, change: dict) -> None:
        """
        Callback function to show or hide schedule dropdown based on checkbox value.

        Args:
            change (dict): Dictionary containing the new value of the schedule checkbox.
        """
        if change["new"]:
            self.show_schedules()
            self.reset_widgets()
            self.hide_bracket_resistor()
        else:
            self.hide_schedules()
            self.reset_widgets()
            self.hide_bracket_resistor()

    def show_schedules(self) -> None:
        """
        Show schedule dropdown and options.
        """
        self.schedule_dropdown.options = [SELECT_OPTION] + [schedule.value for schedule in ScheduleOption]
        self.schedule_dropdown.layout.visibility = "visible"

    def hide_schedules(self) -> None:
        """
        Hide schedule dropdown.
        """
        self.schedule_dropdown.value = SELECT_OPTION
        self.schedule_dropdown.layout.visibility = "hidden"

    def reset_widgets(self) -> None:
        """
        Reset widgets to their initial state.
        """
        self.part_number_input.value = ""
        self.type_dropdown.value = SELECT_TYPE_OPTION
        self.option_dropdown.value = SELECT_OPTION
        self.part_number_input.disabled = self.chbx_schedule.value
        self.type_dropdown.disabled = self.chbx_schedule.value
        self.option_dropdown.disabled = self.chbx_schedule.value
        self.reset_checkboxes(False)

    def execute_automation(self) -> None:
        """
        Execute automation based on user-selected options and inputs.
        """
        if self.chbx_schedule.value:
            if self.schedule_dropdown.value != SELECT_OPTION:
                self.btn_generate_pdf.disabled = True
                self.execute_schedule()
            else:
                print("Please select a valid option for the Schedule Document.")
        else:
            self.process_automation()

    def process_automation(self) -> None:
        """Process automation based on the type and options selected."""
        self.part_number = str(self.part_number_input.value)
        if not self.part_number:
            print("Please enter a part number.")
            return
        
        selected_type =  DropdownType(self.type_dropdown.value)
        selected_option = TypeOption(self.option_dropdown.value)
        if selected_type == SELECT_OPTION:
            print("Please select a valid type.")
            return
        
        if not self.is_valid_part_number(selected_type):
            return
        
        self.execute_based_on_option(selected_option, selected_type)

    def is_valid_part_number(self, selected_type: DropdownType) -> bool:
        """
        Validate the part number based on the selected type.

        Args:
            selected_type (DropdownType): The selected dropdown type.

        Returns:
            bool: True if the part number is valid, False otherwise.
        """
        if selected_type in [DropdownType.FLC_25, DropdownType.FLR_25]:
            self.btn_generate_pdf.disabled = True
            self.letter_identifier = self.utils.validate_part_number(self.part_number)
            if not self.letter_identifier:
                return False
            if self.letter_identifier == LetterIdentifier.N.value:
                print(
                    f"\033[1;38;2;255;0;0m\nUnexpected part number: {self.part_number}. For Fusion 3, part numbers should include 'SC', 'R', or 'X', but not 'N'.\033[0m"
                )
                return False
        elif selected_type in [DropdownType.FLC_20, DropdownType.FLR_21]:
            self.btn_generate_pdf.disabled = True
            self.letter_identifier = self.utils.validate_part_number(self.part_number)
            if not self.letter_identifier:
                return False
        elif selected_type in [
            DropdownType.FLR_25_BRACKET,
            DropdownType.FLC_25_BRACKET,
            DropdownType.FLC_25_COVER,
        ]:
            self.btn_generate_pdf.disabled = True
            if not self.utils.validate_bracket_cover_number(self.part_number):
                print("Invalid bracket or cover part number.")
                return False
        else:
            print("Selected type is not supported.")
            return False
        return True

    def execute_based_on_option(self, option: TypeOption, selected_type: DropdownType) -> None:
        """
        Execute the appropriate automation based on the selected option.

        Args:
            option (TypeOption): The selected dropdown option.
            selected_type (DropdownType): The selected dropdown type.
        """
        option_actions = {
            TypeOption.COMPARE_INI_FILE: self.execute_ini_automation,
            TypeOption.BOM_STRUCTURE: self.execute_bom_automation,
            TypeOption.DOCUMENT_VALIDATION: self.execute_content_automation,
            TypeOption.AFTERMARKET_PARTS: self.execute_aftermarket_automation,
            TypeOption.ATTRIBUTES_VALIDATION: self.execute_attributes_validation,
        }
        if selected_type in [
            DropdownType.FLR_25,
            DropdownType.FLC_25,
            DropdownType.FLC_20,
            DropdownType.FLR_21
        ]:
            action = option_actions.get(option)
            if action:
                action()
            else:
                print("Please select a valid option.")
        elif selected_type in [
            DropdownType.FLR_25_BRACKET,
            DropdownType.FLC_25_BRACKET,
            DropdownType.FLC_25_COVER,
        ]:
            if option in [
                TypeOption.BOM_STRUCTURE,
                TypeOption.DOCUMENT_VALIDATION,
                TypeOption.AFTERMARKET_PARTS,
            ]:
                option_actions.get(option)()
            else:
                print("Please select a valid option for bracket or cover.")
        else:
            print("Selected type is not supported.")

    def main(self) -> None:
        """Main method to interact with the user and display widgets."""
        container_init = widgets.HBox(
            [self.part_number_input, self.chbx_schedule, self.schedule_dropdown]
        )
        container = widgets.HBox(
            [self.type_dropdown, self.chbx_resistor, self.chbx_bracket]
        )
        display(container_init)
        display(container)
        options_buttons_container = widgets.HBox(
            [self.option_dropdown, self.btn_generate_pdf, self.output]
        )
        display(options_buttons_container)

        widgets.interact_manual(self.execute_automation)
