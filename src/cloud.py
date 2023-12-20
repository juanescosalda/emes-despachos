import re
import json
import logging
import utils as utils
import gspread
import gspread.utils as gs_utils
from google.oauth2.service_account import Credentials
from gspread_formatting import *
from typing import List, Dict, Tuple, Any
from datetime import date, datetime


class CloudServices:

    EMPTY_COLS = \
        [
            "Valor cobrado",
            "Mensajero",
            "Hora salida",
            "Observaciones"
        ]

    REPORT_COLS = \
        [
            'Mensajero',
            'Zona',
            'Factura',
            'Cliente',
            'Valor cobrado',
            'Forma de pago',
            'Observaciones',
            'Notas Tesorería'
        ]

    def __init__(
            self,
            folder_id: str,
            json_path: str,
            zones: List[str]):
        """
        Create Google auth and Google Sheets objects
        """
        scopes = \
            [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]

        credentials = Credentials.from_service_account_file(
            json_path,
            scopes=scopes
        )

        # Initialize sheets list
        self.__sheets = []
        self.__names = {}

        # Zones (constant attribute)
        self.ZONES = zones

        # Create client object
        self.client = gspread.authorize(credentials)

        # Initialize cloud services
        self._folder_id = folder_id

        # Get general settings
        f = open('config/settings.json')

        settings = json.load(f)

        self._column_settings = settings["columns"]

        # Initial sheet dimensions
        self.MAX_ROWS = 500
        self.MAX_COLS = len(self._column_settings)

        self.COLUMN_NAMES = \
            [
                value["name"] for value in self._column_settings
            ]

        # Set last col letter
        self._max_col = self.__get_last_col()

        # Set all files in folder
        self._all_files = self.client.list_spreadsheet_files(
            folder_id=folder_id)

    @property
    def sheets(self) -> List:
        return self.__sheets

    @property
    def folder_id(self) -> List:
        return self._folder_id

    @property
    def names(self) -> Dict[str, List[str]]:
        return self.__names

    @names.setter
    def names(self, in_names: Dict[str, List[str]]) -> None:
        self.__names = in_names

    def init(self, names: Dict[str, List[str]]) -> None:
        """
        Create Spreadsheet
        """
        # Get couriers and packers names
        self.__names = names

        # Initialize the necessary variables and objects
        self.__initialize()

    def __get_last_col(self) -> str:
        """
        Get last column as letter

        Returns:
            str: Column name
        """
        col = gs_utils.rowcol_to_a1(
            row=1,
            col=self.MAX_COLS
        )

        return re.split(r'(\d+|[^\d]+)', col)[1]

    def __get_color(
            self,
            red: int,
            blue: int,
            green: int) -> Color:
        """
        Get object color based on RGB code
        """
        return Color(
            utils.rgb2p(red),
            utils.rgb2p(blue),
            utils.rgb2p(green)
        )

    def __create_format_cell(
            self,
            bg_rgb: Tuple[int, int, int],
            fg_rgb: Tuple[int, int, int],
            bold: bool,
            h_align: str = 'CENTER',
            v_align: str = 'MIDDLE',
            number_format: NumberFormat | None = None) -> Dict:
        """
        Create cell format

        Args:
            bg_rgb (Tuple): RGC background color
            fg_rgb (Tuple): RGC foreground color
            bold (bool): True if bold
            h_align (str, optional): Horizontal alignment. Defaults to 'CENTER'.
            v_align (str, optional): Vertical alignment. Defaults to 'MIDDLE'.
            number_format (NumberFormat, optional): Numerical format. Defaults to None.

        Returns:
            Dict: Specified CellFormat converted to dict
        """
        bg_clr = self.__get_color(*bg_rgb)
        fg_clr = self.__get_color(*fg_rgb)

        return CellFormat(
            backgroundColor=bg_clr,
            textFormat=TextFormat(
                bold=bold,
                foregroundColor=fg_clr,
            ),
            horizontalAlignment=h_align,
            verticalAlignment=v_align,
            numberFormat=number_format
        ).to_props()

    def __set_formats_specs(self) -> Dict:
        """
        Set cells formats
        """
        return \
            {
                "header": self.__create_format_cell(
                    bg_rgb=(0, 150, 214),
                    fg_rgb=(250, 250, 250),
                    bold=True
                ),
                "center": self.__create_format_cell(
                    bg_rgb=(255, 255, 255),
                    fg_rgb=(69, 69, 69),
                    bold=False
                ),
                "left": self.__create_format_cell(
                    bg_rgb=(255, 255, 255),
                    fg_rgb=(69, 69, 69),
                    bold=False,
                    h_align='LEFT'
                ),
                "currency": self.__create_format_cell(
                    bg_rgb=(255, 255, 255),
                    fg_rgb=(69, 69, 69),
                    bold=False,
                    number_format=NumberFormat(
                        type='CURRENCY',
                        pattern='$ #,###'
                    )
                )
            }

    def __initialize(self) -> None:
        """
        Create Workbook and Worksheets objects
        """
        try:
            filename = date.today().strftime("%Y-%m-%d")

            # Create CellFormat objects
            self._formats = self.__set_formats_specs()

            # Create Google Sheets workbook if not exist
            if self.__exist_file(filename):
                self.workbook = self.select_file(filename)
            else:
                self.workbook = self.client.create(
                    title=filename,
                    folder_id=self._folder_id
                )

                # Create all sheets
                self.__create_all_sheets()

        except Exception as e:
            logging.error(
                f"Error al inicializar la hoja de Google Sheets >>> {e}",
                exc_info=True
            )

    def __create_sheet(self, zone: str) -> gspread.Worksheet:
        """
        Args:
            zone (str): order region

        Returns:
            Worksheet: created gspread Worksheet
        """
        sheet = self.workbook.add_worksheet(
            title=zone,
            rows=self.MAX_ROWS,
            cols=self.MAX_COLS
        )

        sheet.clear()

        return sheet

    def __create_all_sheets(self) -> None:
        """
        Create five region sheets
        """
        try:
            self.__sheets = [
                self.__create_sheet(zone) for zone in self.ZONES
            ]

            # Get default sheet
            default_sheet = self.workbook.worksheet('Sheet1')

            # Delete default workbook sheet
            self.workbook.del_worksheet(default_sheet)

            # Format new created sheets
            self.__format_all_sheets()

        except (KeyError, AttributeError, ValueError):
            logging.error(
                'Error creando las hojas del archivo de Google Sheets',
                exc_info=True
            )

    def __set_formats(
            self,
            sheet: gspread.Worksheet) -> Dict[str, List]:
        """
        Set columns width

        Args:
            sheet (gspread.Worksheet): Sheet object

        Returns:
            Dict[str, List]: Response batch_update method
        """
        body = \
            {
                'requests': [
                    {
                        'repeatCell': {
                            'range': {
                                'sheetId': sheet.id,
                                'startRowIndex': 0,
                                'endRowIndex': 1,
                                'startColumnIndex': 0,
                                'endColumnIndex': self.MAX_COLS
                            },
                            'cell': {
                                'userEnteredFormat': self._formats["header"]
                            },
                            'fields': 'userEnteredFormat'
                        }
                    }
                ]
            }

        for idx, value in enumerate(self._column_settings):
            body['requests'].extend(
                [
                    {
                        'updateDimensionProperties': {
                            'range': {
                                'sheetId': sheet.id,
                                'dimension': 'COLUMNS',
                                'startIndex': idx,
                                'endIndex': idx + 1
                            },
                            'properties': {
                                'pixelSize': value["width"]
                            },
                            'fields': 'pixelSize'
                        }
                    },
                    {
                        'repeatCell': {
                            'range': {
                                'sheetId': sheet.id,
                                'startRowIndex': 1,
                                'endRowIndex': self.MAX_ROWS,
                                'startColumnIndex': idx,
                                'endColumnIndex': idx + 1
                            },
                            'cell': {
                                'userEnteredFormat': self._formats[value["format"]]
                            },
                            'fields': 'userEnteredFormat'
                        }
                    }
                ]
            )

        # Set common column formats
        packers_idx = self.get_col_idx("Empacador")
        couriers_idx = self.get_col_idx("Mensajero")

        body['requests'].extend(
            [
                {
                    'repeatCell': {
                        'range': {
                            'sheetId': sheet.id,
                            'startRowIndex': 0,
                            'endRowIndex': 1,
                            'startColumnIndex': 0,
                            'endColumnIndex': self.MAX_COLS
                        },
                        'cell': {
                            'userEnteredFormat': self._formats["header"]
                        },
                        'fields': 'userEnteredFormat'
                    }
                },
                {
                    'updateDimensionProperties': {
                        'range': {
                            'sheetId': sheet.id,
                            'dimension': 'ROWS',
                            'startIndex': 0,
                            'endIndex': self.MAX_ROWS
                        },
                        'properties': {
                            'pixelSize': 28
                        },
                        'fields': 'pixelSize'
                    }
                },
                {
                    "setDataValidation": {
                        "range": {
                            "sheetId": sheet.id,
                            "startRowIndex": 1,
                            "endRowIndex": self.MAX_ROWS,
                            "startColumnIndex": couriers_idx,
                            "endColumnIndex": couriers_idx + 1
                        },
                        "rule": {
                            "condition": {
                                "type": 'ONE_OF_LIST',
                                "values": [
                                        {"userEnteredValue": name} for name in self.__names["couriers"]
                                ],
                            },
                            "showCustomUi": True,
                            "strict": True
                        }
                    }
                },
                {
                    "setDataValidation": {
                        "range": {
                            "sheetId": sheet.id,
                            "startRowIndex": 1,
                            "endRowIndex": self.MAX_ROWS,
                            "startColumnIndex": packers_idx,
                            "endColumnIndex": packers_idx + 1
                        },
                        "rule": {
                            "condition": {
                                "type": 'ONE_OF_LIST',
                                "values": [
                                        {"userEnteredValue": name} for name in self.__names["packers"]
                                ],
                            },
                            "showCustomUi": True,
                            "strict": True
                        }
                    }
                }
            ]
        )

        res = self.workbook.batch_update(body)

        return res

    def __format_sheet(
            self,
            sheet: gspread.Worksheet) -> None:
        """
        Apply format to each sheet

        sheet (gspread.Worksheet): sheet to format
        """
        # Insert titles
        self.insert_row(self.COLUMN_NAMES, sheet.title)

        # Set column formats
        self.__set_formats(sheet)

    def __format_all_sheets(self) -> None:
        """
        Apply format to all sheets
        """
        try:
            for sheet in self.__sheets:
                self.__format_sheet(sheet)

        except (ValueError, KeyError, AttributeError, TypeError):
            logging.error(
                'Error al agregar formato a todas las hojas e insertando títulos',
                exc_info=True
            )

    def get_col_idx(self, col_name: str) -> int:
        """
        Returns column index based on column name

        Args:
            col_name (str): Column name

        Returns:
            int: Column index
        """
        col_idx = self.COLUMN_NAMES.index(col_name)

        return col_idx

    def get_total_rows(self, sheet: gspread.Worksheet) -> int:
        """
        Get total rows of specified sheet

        Args:
            sheet (gspread.Worksheet): Zone sheet

        Returns:
            int: Number of rows
        """
        rows = len(sheet.col_values(1)) + 1

        return rows

    def update_users(self, names: Dict[str, List[str]]) -> Any:
        """
        Update "Mensajeros" y "Empacadores"
        """
        # Update user names member varible
        self.__names = names

        # Get column indexes
        packers_idx = self.get_col_idx("Empacador")
        couriers_idx = self.get_col_idx("Mensajero")

        # Request dict format
        body = {"requests": []}

        for zone in self.ZONES:
            sheet = self.workbook.worksheet(zone)

            body["requests"].extend(
                [
                    {
                        "setDataValidation": {
                            "range": {
                                "sheetId": sheet.id,
                                "startRowIndex": 1,
                                "endRowIndex": self.MAX_ROWS,
                                "startColumnIndex": couriers_idx,
                                "endColumnIndex": couriers_idx + 1
                            },
                            "rule": {
                                "condition": {
                                    "type": 'ONE_OF_LIST',
                                    "values": [
                                        {"userEnteredValue": name} for name in self.__names["couriers"]
                                    ],
                                },
                                "showCustomUi": True,
                                "strict": True
                            }
                        }
                    },
                    {
                        "setDataValidation": {
                            "range": {
                                "sheetId": sheet.id,
                                "startRowIndex": 1,
                                "endRowIndex": self.MAX_ROWS,
                                "startColumnIndex": packers_idx,
                                "endColumnIndex": packers_idx + 1
                            },
                            "rule": {
                                "condition": {
                                    "type": 'ONE_OF_LIST',
                                    "values": [
                                        {"userEnteredValue": name} for name in self.__names["packers"]
                                    ],
                                },
                                "showCustomUi": True,
                                "strict": True
                            }
                        }
                    }
                ]
            )

        res = self.workbook.batch_update(body)

        return res

    def insert_row(
            self,
            values: List[Any],
            zone: str) -> None:
        """
        Insert element into last row

        Args:
            values (List[Any]): Customer invoice info
            zone (str): Zone name (also Sheet title)
        """
        try:
            sheet = self.workbook.worksheet(zone)

            curr_row = self.get_total_rows(sheet)

            # Update current non written row
            sheet.update(
                f'A{curr_row}:{self._max_col}{curr_row}',
                [values]
            )

            # Include basic filter in the header
            sheet.set_basic_filter(
                f'A1:{self._max_col}{curr_row}'
            )

        except Exception as e:
            logging.error(
                f"Error {e} al insertar nueva fila al archivo de Google Sheets",
                exc_info=True
            )

    def convert_to_sheet_data(
            self,
            rows_data: List[List[Any]]) -> Dict[str, List[List[Any]]]:
        """
        Converts UI table widget table data into GS required format
        """
        current_time = datetime.now().strftime(
            "%Y-%m-%d %H:%M")

        values = \
            [
                [
                    value[3],
                    value[1],
                    value[0],
                    value[9],
                    utils.safe_int(value[6]),
                    utils.safe_int(value[7]),
                    utils.safe_int(value[8]),
                    value[4],
                    utils.from_currency(value[5]),
                    0,
                    value[10],
                    value[11],
                    value[2],
                    current_time,
                    ""
                ]
                for value in rows_data
            ]

        out_data = \
            {
                zone: [v[1:] for v in values if v[0] == zone]
                for zone in self.ZONES
            }

        return out_data

    def insert_rows(
            self,
            rows_data: List[List[Any]]) -> bool:
        """
        Insert element into last row

        Args:
            rows_data (List[Any]): Customer invoice info
        """
        try:
            dict_data = self.convert_to_sheet_data(rows_data)

            for zone, value in dict_data.items():
                if value:
                    # Get sheet object
                    sheet = self.workbook.worksheet(zone)

                    # Get current row
                    curr_row = self.get_total_rows(sheet)

                    # Get max row
                    max_row = curr_row + len(value) - 1
                    range_str = f'A{curr_row}:{self._max_col}{max_row}'

                    # Update current non written row
                    res = sheet.update(range_str, value)

                    if 'updatedRange' in res:
                        # Include basic filter in the header
                        sheet.set_basic_filter(
                            f'A1:{self._max_col}{max_row}'
                        )
                    else:
                        logging.warning(
                            f"Error al actualizar los datos en la hoja de Google Sheets ({res})"
                        )
                        return False

        except Exception as e:
            logging.error(
                f"Error {e} al insertar nueva fila al archivo de Google Sheets",
                exc_info=True
            )
            return False

        return True

    def __find_row_by_query(
            self,
            sheet: gspread.Worksheet,
            query: Any,
            col_name: str) -> int:
        """
        Find row index by column name

        Args:
            sheet (gspread.Worksheet): Worksheet
            query (Any): Column item to find
            col_name (str): Header column name

        Returns:
            int: Row index
        """
        col_idx = self.get_col_idx(col_name)

        cell = sheet.find(
            query,
            in_column=col_idx + 1
        )

        return cell.row if cell else -1

    def update_row(
            self,
            zone: str,
            query: Any,
            values: List[Any]) -> None:
        """
        Update row if there are any changes

        Args:
            zone (str): Zone in database
            query (Any): File query
            values (List): 2D list with row values
        """
        sheet = self.workbook.worksheet(zone)

        row = self.__find_row_by_query(sheet, query, "Factura")

        if row != -1:
            range_name = f'A{row}:{self._max_col}{row}'
            row_values = sheet.row_values(row)

            idx = self.get_col_idx("Valor cobrado")
            values[0][idx] = row_values[idx]

            idx = self.get_col_idx("Mensajero")
            values[0][idx] = row_values[idx]

            idx = self.get_col_idx("Hora salida")
            values[0][idx] = row_values[idx]

            idx = self.get_col_idx("Observaciones")
            values[0][idx] = row_values[idx]

            # Update sheet data
            sheet.update(range_name, values)

    def delete_row(self, query: Any) -> List[Any]:
        """
        Delete row of specified query

        Args:
            query (Any): Column item to find

        Returns:
            List[Any]: Deleted values
        """
        for zone in self.ZONES:
            sheet = self.workbook.worksheet(zone)
            row = self.__find_row_by_query(sheet, query, "Factura")

            if row != -1:
                break

        # Get values previous to delete
        deleted_values = sheet.row_values(row)

        # Delete row by query
        sheet.delete_rows(row)

        return deleted_values

    def __get_sheet_names(
            self,
            workbook: gspread.Spreadsheet) -> List[str]:
        """
        Get names of all sheets of current Spreadsheet

        Args:
            workbook (gspread.Spreadsheet): Current Spreadsheet

        Returns:
            List[str]: Names of sheets if workbook exists
        """
        sheets = []

        if workbook:
            sheets = [ws.title for ws in workbook.worksheets()]

        return sheets

    def __get_all_values(self) -> List[List]:
        """
        Get values of all sheets in one list

        Returns:
            List[List]: contains all rows of all sheets
        """
        return sum(
            [
                self.workbook.worksheet(zone).get_all_values()[1:]
                for zone in self.ZONES
            ],
            []
        )

    def __exist_sheet(
            self,
            workbook: gspread.Spreadsheet,
            sheet_name: str) -> bool:
        """
        Check if sheet exists

        Args:
            workbook (gspread.Spreadsheet): Current spreadsheet
            sheet_name (str): Sheet name

        Returns:
            bool: True if sheet exists
        """
        sheet_names = self.__get_sheet_names(workbook)

        return sheet_name in sheet_names

    def __exist_file(self, file: str) -> bool:
        """
        Checks if file exists in specified folder

        Args:
            file (str): File name

        Returns:
            bool: True if there is a file in the folder
        """
        dict_files = self.client.list_spreadsheet_files(
            folder_id=self._folder_id)

        files = [d['name'] for d in dict_files]

        return file in files

    def select_file(self, filename: str) -> gspread.Spreadsheet:
        """
        Select the file specified by `filename` from Google Drive.

        Args:
            filename (str): Name of the file to select

        Returns:
            gspread.Spreadsheet: The selected file as a `gspread.Spreadsheet` object
        """
        try:
            ss = self.client.open(filename, self._folder_id)
            return ss

        except Exception as e:
            logging.error(
                f'Error opening specified filename: {e}',
                exc_info=True
            )

    def __create_report_sheet(
            self,
            filename: str,
            name: str,
            headers: List[str],
            reset: bool) -> Tuple[gspread.Spreadsheet, gspread.Worksheet, bool]:
        """
        Create report sheet

        Args:
            filename (str): Name of file to get report
            name (str): Sheet name
            headers (List[str]): List of headers for the sheet
            reset (bool): True if reset sheet (only headers retained)

        Returns:
            Tuple: Contains the Spreadsheet object, the Worksheet object, and a bool flag indicating if the sheet existed.
        """
        try:
            workbook = self.select_file(filename)
            sheet_exists = self.__exist_sheet(workbook, name)

            if sheet_exists:
                report_sheet = workbook.worksheet(name)

                prev_values = report_sheet.get_all_values()[1:]

                if reset:
                    report_sheet.resize(rows=1)
            else:
                report_sheet = workbook.add_worksheet(
                    title=name,
                    rows=1,
                    cols=len(headers)
                )

                # Insert headers
                report_sheet.append_row(
                    values=headers
                )

                prev_values = list()

            return workbook, report_sheet, sheet_exists, prev_values

        except Exception as e:
            logging.error(
                f'Error creating report sheet: {e}',
                exc_info=True
            )

            return None, None, False, None

    def get_client_name(self, client: str) -> str:
        """
        Args:
            client (str): Client name with address and name

        Returns:
            str: Client name
        """
        separator = '-'

        if client.find(separator) != -1:
            parts = client.split(separator)
            return parts[0].strip()
        else:
            return client

    def merge_sheet_values(
            self,
            ss: gspread.Spreadsheet,
            code_values: List[Any],
            prev_codes: List[Any]) -> List[List]:
        """
        Unifies the values of all sheets into a single one

        Args:
            ss (gspread.Spreadsheet): Current Spreadsheet
            code_values (List): Code column values
            prev_codes (List): Prev modified codes

        Returns:
            List[List]: Sheets merged values 
        """
        values_to_append = []

        # Get column index
        code_idx = self.get_col_idx("Factura")
        client_idx = self.get_col_idx("Cliente")
        msg_idx = self.get_col_idx("Mensajero")
        value_idx = self.get_col_idx("Valor cobrado")
        forma_idx = self.get_col_idx("Forma de pago")
        obs_idx = self.get_col_idx("Observaciones")

        for zone in self.ZONES:
            sheet = ss.worksheet(zone)
            rows = sheet.get_all_values()

            for row in rows[1:]:
                code = row[code_idx]

                if code == 0 or code not in code_values[1:]:
                    client = self.get_client_name(row[client_idx])
                    checked = utils.str_to_bool(
                        prev_codes[code]) if code in prev_codes else False

                    if row[msg_idx] != "":
                        values = \
                            [
                                row[msg_idx],
                                sheet.title,
                                code,
                                client,
                                utils.str_to_int(row[value_idx]),
                                row[forma_idx],
                                row[obs_idx],
                                checked
                            ]

                        values_to_append.append(values)

        return values_to_append

    def __create_common_report(
            self,
            name: str,
            filename: str,
            emails: List[str]) -> Dict:
        """
        Common format to create Dispatches or Treasury format

        Args:
            name (str): Report name ('treasury' or 'dispatches')
            filename (str): Name of the file to get the report
            emails (List[str]): Email accounts for whom the sheet is non-protected

        Returns:
            dict: Contains batch update response
        """
        workbook, report_sheet, sheet_exists, prev_values = \
            self.__create_report_sheet(
                filename=filename,
                name=name,
                headers=CloudServices.REPORT_COLS,
                reset=True
            )

        prev_codes = \
            {v[2]: v[7] for v in prev_values} if sheet_exists else {}

        if report_sheet is None:
            return {}

        code_values = report_sheet.col_values(2)
        start_row = len(code_values)

        # Merge all sheets data into one 2D list
        all_values = self.merge_sheet_values(
            workbook,
            code_values,
            prev_codes
        )

        if not all_values:
            return {}

        report_sheet.append_rows(
            values=all_values
        )

        # Get max dimensions
        max_row = len(report_sheet.col_values(1))
        max_col = len(CloudServices.REPORT_COLS)

        # Create request dict
        body = {'requests': []}

        if name == "Despachos":
            body['requests'].append(
                {
                    'updateDimensionProperties': {
                        'range': {
                            'sheetId': report_sheet.id,
                            'dimension': 'COLUMNS',
                            'startIndex': 7,
                            'endIndex': max_col
                        },
                        "properties": {
                            "hiddenByUser": True,
                        },
                        "fields": 'hiddenByUser'
                    }
                }
            )

        # Request used when creating the sheet for the first time
        if not sheet_exists:
            body['requests'].extend(
                [
                    {
                        "addProtectedRange": {
                            "protectedRange": {
                                "range": {
                                    "sheetId": report_sheet.id,
                                },
                                "editors": {
                                    "domainUsersCanEdit": False,
                                    "users": emails
                                },
                                "warningOnly": False
                            }
                        }
                    }
                ]
            )

        # Add request for updating row sizes
        body['requests'].extend(
            [
                {
                    'updateDimensionProperties': {
                        'range': {
                            'sheetId': report_sheet.id,
                            'dimension': 'COLUMNS',
                            'startIndex': 0,
                            'endIndex': 1
                        },
                        'properties': {
                            'pixelSize': 180
                        },
                        'fields': 'pixelSize'
                    }
                },
                {
                    'updateDimensionProperties': {
                        'range': {
                            'sheetId': report_sheet.id,
                            'dimension': 'COLUMNS',
                            'startIndex': 1,
                            'endIndex': 3
                        },
                        'properties': {
                            'pixelSize': 90
                        },
                        'fields': 'pixelSize'
                    }
                },
                {
                    'updateDimensionProperties': {
                        'range': {
                            'sheetId': report_sheet.id,
                            'dimension': 'COLUMNS',
                            'startIndex': 3,
                            'endIndex': 4
                        },
                        'properties': {
                            'pixelSize': 300
                        },
                        'fields': 'pixelSize'
                    }
                },
                {
                    'updateDimensionProperties': {
                        'range': {
                            'sheetId': report_sheet.id,
                            'dimension': 'COLUMNS',
                            'startIndex': 4,
                            'endIndex': 5
                        },
                        'properties': {
                            'pixelSize': 120
                        },
                        'fields': 'pixelSize'
                    }
                },
                {
                    'updateDimensionProperties': {
                        'range': {
                            'sheetId': report_sheet.id,
                            'dimension': 'COLUMNS',
                            'startIndex': 5,
                            'endIndex': 6
                        },
                        'properties': {
                            'pixelSize': 160
                        },
                        'fields': 'pixelSize'
                    }
                },
                {
                    'updateDimensionProperties': {
                        'range': {
                            'sheetId': report_sheet.id,
                            'dimension': 'COLUMNS',
                            'startIndex': 6,
                            'endIndex': 7
                        },
                        'properties': {
                            'pixelSize': 250
                        },
                        'fields': 'pixelSize'
                    }
                },
                {
                    'updateDimensionProperties': {
                        'range': {
                            'sheetId': report_sheet.id,
                            'dimension': 'COLUMNS',
                            'startIndex': 7,
                            'endIndex': 8
                        },
                        'properties': {
                            'pixelSize': 130
                        },
                        'fields': 'pixelSize'
                    }
                },
                {
                    'updateDimensionProperties': {
                        'range': {
                            'sheetId': report_sheet.id,
                            'dimension': 'ROWS',
                            'startIndex': start_row - 1,
                            'endIndex': max_row
                        },
                        'properties': {
                            'pixelSize': 30
                        },
                        'fields': 'pixelSize'
                    }
                },
                {
                    'setBasicFilter': {
                        'filter': {
                            'range': {
                                'sheetId': report_sheet.id,
                                'startRowIndex': 0,
                                'endRowIndex': max_row,
                                'startColumnIndex': 0,
                                'endColumnIndex': max_col
                            },
                        }
                    }
                },
                {
                    'repeatCell': {
                        'range': {
                            'sheetId': report_sheet.id,
                            'startRowIndex': 0,
                            'endRowIndex': 1,
                            'startColumnIndex': 0,
                            'endColumnIndex': max_col
                        },
                        'cell': {
                            'userEnteredFormat': self._formats["header"]
                        },
                        'fields': 'userEnteredFormat'
                    }
                },
                {
                    'repeatCell': {
                        'range': {
                            'sheetId': report_sheet.id,
                            'startRowIndex': 1,
                            'endRowIndex': max_row,
                            'startColumnIndex': 0,
                            'endColumnIndex': 1
                        },
                        'cell': {
                            'userEnteredFormat': self._formats["left"]
                        },
                        'fields': 'userEnteredFormat'
                    }
                },
                {
                    'repeatCell': {
                        'range': {
                            'sheetId': report_sheet.id,
                            'startRowIndex': 1,
                            'endRowIndex': max_row,
                            'startColumnIndex': 1,
                            'endColumnIndex': 3
                        },
                        'cell': {
                            'userEnteredFormat': self._formats["center"]
                        },
                        'fields': 'userEnteredFormat'
                    }
                },
                {
                    'repeatCell': {
                        'range': {
                            'sheetId': report_sheet.id,
                            'startRowIndex': 1,
                            'endRowIndex': max_row,
                            'startColumnIndex': 3,
                            'endColumnIndex': 4
                        },
                        'cell': {
                            'userEnteredFormat': self._formats["left"]
                        },
                        'fields': 'userEnteredFormat'
                    }
                },
                {
                    'repeatCell': {
                        'range': {
                            'sheetId': report_sheet.id,
                            'startRowIndex': 1,
                            'endRowIndex': max_row,
                            'startColumnIndex': 4,
                            'endColumnIndex': 5
                        },
                        'cell': {
                            'userEnteredFormat': self._formats["currency"]
                        },
                        'fields': 'userEnteredFormat'
                    }
                },
                {
                    'repeatCell': {
                        'range': {
                            'sheetId': report_sheet.id,
                            'startRowIndex': 1,
                            'endRowIndex': max_row,
                            'startColumnIndex': 5,
                            'endColumnIndex': max_col
                        },
                        'cell': {
                            'userEnteredFormat': self._formats["center"]
                        },
                        'fields': 'userEnteredFormat'
                    }
                },
                {
                    "repeatCell": {
                        "cell": {
                            "dataValidation": {
                                "condition": {"type": "BOOLEAN"}
                            }
                        },
                        "range": {
                            "sheetId": report_sheet.id,
                            "startRowIndex": 1,
                            "endRowIndex": max_row,
                            "startColumnIndex": max_col - 1,
                            "endColumnIndex": max_col
                        },
                        "fields": "dataValidation"
                    }
                }
            ]
        )

        res = workbook.batch_update(body)

        return res

    def create_treasury_report(self, filename: str) -> Dict:
        """
        Args:
            filename (str): File name (date based)

        Returns:
            Dict: Batch update response
        """
        self.__create_common_report(
            name='Tesorería',
            filename=filename,
            emails=[
                "depositoemestesoreria1@gmail.com",
                "depositoemescontabilidad1@gmail.com"
            ]
        )

    def create_dispatches_report(self, filename: str) -> Dict:
        """
        Args:
            filename (str): File name (date based)

        Returns:
            Dict: Batch update response
        """
        self.__create_common_report(
            name='Despachos',
            filename=filename,
            emails=[
                "depositoemesdespachos1@gmail.com"
            ]
        )

    def create_couriers_report(self, filename: str) -> Dict:
        """
        Create couriers (messengers) daily report

        Args:
            filename (str): Name of the file to get the report

        Returns:
            dict: Contains batch update response
        """
        workbook, report_sheet, sheet_exists = self.__create_report_sheet(
            filename=filename,
            name='Mensajeros',
            headers=[
                'Mensajero',
                'Valor cobrado',
                'Transportes',
                'Cajas',
                'Bolsas',
                'Lios'
            ],
            reset=True
        )

        if report_sheet is None:
            return {}

        # store data
        data = {
            name:
            {
                "value": 0,
                "count": 0,
                "cajas": 0,
                "bolsas": 0,
                "lios": 0
            } for name in self.__names["couriers"]
        }

        rows = self.__get_all_values()

        # iterate over each row, and get the values
        for row in rows:
            name = row[9]
            if name != '':
                data[name]['value'] += utils.str_to_int(row[6])
                data[name]['count'] += 1
                data[name]['cajas'] += utils.safe_int(row[2])
                data[name]['bolsas'] += utils.safe_int(row[3])
                data[name]['lios'] += utils.safe_int(row[4])

        # append multiple rows at once
        report_sheet.append_rows(
            values=[[k] + list(v.values()) for k, v in data.items()]
        )

        # get dimensiones
        max_row = len(report_sheet.col_values(1))
        max_col = len(data.keys()) + 1

        # create requests
        body = {'requests': []}

        # protect sheet
        if not sheet_exists:
            body['requests'].append({
                "addProtectedRange": {
                    "protectedRange": {
                        "range": {
                            "sheetId": report_sheet.id,
                        },
                        "editors": {
                            "domainUsersCanEdit": False,
                            "users": [
                                'depositoemesdespachos1@gmail.com',
                                'jigomez6025@gmail.com'
                            ]
                        },
                        "warningOnly": False
                    }
                }
            })

        body['requests'].append({
            'updateDimensionProperties': {
                'range': {
                    'sheetId': report_sheet.id,
                    'dimension': 'COLUMNS',
                    'startIndex': 0,
                    'endIndex': 1
                },
                'properties': {
                    'pixelSize': 250
                },
                'fields': 'pixelSize'
            }
        })

        body['requests'].append({
            'updateDimensionProperties': {
                'range': {
                    'sheetId': report_sheet.id,
                    'dimension': 'COLUMNS',
                    'startIndex': 1,
                    'endIndex': 2
                },
                'properties': {
                    'pixelSize': 140
                },
                'fields': 'pixelSize'
            }
        })

        body['requests'].append({
            'updateDimensionProperties': {
                'range': {
                    'sheetId': report_sheet.id,
                    'dimension': 'COLUMNS',
                    'startIndex': 2,
                    'endIndex': max_col
                },
                'properties': {
                    'pixelSize': 100
                },
                'fields': 'pixelSize'
            }
        })

        body['requests'].append({
            'updateDimensionProperties': {
                'range': {
                    'sheetId': report_sheet.id,
                    'dimension': 'ROWS',
                    'startIndex': 0,
                    'endIndex': max_row
                },
                'properties': {
                    'pixelSize': 30
                },
                'fields': 'pixelSize'
            }
        })

        # add basic filter
        body['requests'].append({
            'setBasicFilter': {
                'filter': {
                    'range': {
                        'sheetId': report_sheet.id,
                        'startRowIndex': 0,
                        'endRowIndex': max_row,
                        'startColumnIndex': 0,
                        'endColumnIndex': max_col
                    },
                }
            }
        })

        # add request to format header cells
        body['requests'].append({
            'repeatCell': {
                'range': {
                    'sheetId': report_sheet.id,
                    'startRowIndex': 0,
                    'endRowIndex': 1,
                    'startColumnIndex': 0,
                    'endColumnIndex': max_col
                },
                'cell': {
                    'userEnteredFormat': self._formats["header"]
                },
                'fields': 'userEnteredFormat'
            }
        })

        # add request to format body left cells
        body['requests'].append({
            'repeatCell': {
                'range': {
                    'sheetId': report_sheet.id,
                    'startRowIndex': 1,
                    'endRowIndex': max_row,
                    'startColumnIndex': 0,
                    'endColumnIndex': 1
                },
                'cell': {
                    'userEnteredFormat': self._formats["left"]
                },
                'fields': 'userEnteredFormat'
            }
        })

        # add request to format body currency cells
        body['requests'].append({
            'repeatCell': {
                'range': {
                    'sheetId': report_sheet.id,
                    'startRowIndex': 1,
                    'endRowIndex': max_row,
                    'startColumnIndex': 1,
                    'endColumnIndex': 2
                },
                'cell': {
                    'userEnteredFormat': self._formats["currency"]
                },
                'fields': 'userEnteredFormat'
            }
        })

        # add request to format body cells
        body['requests'].append({
            'repeatCell': {
                'range': {
                    'sheetId': report_sheet.id,
                    'startRowIndex': 1,
                    'endRowIndex': max_row,
                    'startColumnIndex': 2,
                    'endColumnIndex': max_col
                },
                'cell': {
                    'userEnteredFormat': self._formats["center"]
                },
                'fields': 'userEnteredFormat'
            }
        })

        res = workbook.batch_update(body)

        return res

    def create_packers_report(self, filename: str) -> Dict:
        """
        Create packers report

        Args:
            filename (str): Name of the file to get the report

        Returns:
            Dict: Contains batch update response
        """
        workbook, report_sheet, sheet_exists = self.__create_report_sheet(
            filename=filename,
            name='Empacadores',
            headers=[
                'Empacador',
                '# Empaques',
                'Cajas',
                'Bolsas',
                'Lios'
            ],
            reset=True
        )

        if report_sheet is None:
            return {}

        # Store data
        data = {
            name:
            {
                "count": 0,
                "cajas": 0,
                "bolsas": 0,
                "lios": 0
            } for name in self.__names["packers"] + ['Otro']
        }

        rows = self.__get_all_values()

        for row in rows:
            name = row[7] if row[7] != '' else 'Otro'
            if name != '':
                data[name]['count'] += 1
                data[name]['cajas'] += utils.safe_int(row[2])
                data[name]['bolsas'] += utils.safe_int(row[3])
                data[name]['lios'] += utils.safe_int(row[4])

        report_sheet.append_rows(
            values=[[k] + list(v.values()) for k, v in data.items()]
        )

        # Get dimensiones
        max_row = len(report_sheet.col_values(1))
        max_col = len(data.keys()) + 1

        # Create requests
        body = {'requests': []}

        # protect sheet
        if not sheet_exists:
            body['requests'].append({
                "addProtectedRange": {
                    "protectedRange": {
                        "range": {
                            "sheetId": report_sheet.id,
                        },
                        "editors": {
                            "domainUsersCanEdit": False,
                            "users": [
                                'depositoemesdespachos1@gmail.com',
                                'jigomez6025@gmail.com'
                            ]
                        },
                        "warningOnly": False
                    }
                }
            })

        body['requests'].append({
            'updateDimensionProperties': {
                'range': {
                    'sheetId': report_sheet.id,
                    'dimension': 'COLUMNS',
                    'startIndex': 0,
                    'endIndex': 1
                },
                'properties': {
                    'pixelSize': 250
                },
                'fields': 'pixelSize'
            }
        })

        body['requests'].append({
            'updateDimensionProperties': {
                'range': {
                    'sheetId': report_sheet.id,
                    'dimension': 'COLUMNS',
                    'startIndex': 1,
                    'endIndex': 2
                },
                'properties': {
                    'pixelSize': 140
                },
                'fields': 'pixelSize'
            }
        })

        body['requests'].append({
            'updateDimensionProperties': {
                'range': {
                    'sheetId': report_sheet.id,
                    'dimension': 'COLUMNS',
                    'startIndex': 2,
                    'endIndex': max_col
                },
                'properties': {
                    'pixelSize': 100
                },
                'fields': 'pixelSize'
            }
        })

        body['requests'].append({
            'updateDimensionProperties': {
                'range': {
                    'sheetId': report_sheet.id,
                    'dimension': 'ROWS',
                    'startIndex': 0,
                    'endIndex': max_row
                },
                'properties': {
                    'pixelSize': 30
                },
                'fields': 'pixelSize'
            }
        })

        body['requests'].append({
            'setBasicFilter': {
                'filter': {
                    'range': {
                        'sheetId': report_sheet.id,
                        'startRowIndex': 0,
                        'endRowIndex': max_row,
                        'startColumnIndex': 0,
                        'endColumnIndex': max_col
                    },
                }
            }
        })

        # Add request to format header cells
        body['requests'].append({
            'repeatCell': {
                'range': {
                    'sheetId': report_sheet.id,
                    'startRowIndex': 0,
                    'endRowIndex': 1,
                    'startColumnIndex': 0,
                    'endColumnIndex': max_col
                },
                'cell': {
                    'userEnteredFormat': self._formats["header"]
                },
                'fields': 'userEnteredFormat'
            }
        })

        # add request to format body left cells
        body['requests'].append({
            'repeatCell': {
                'range': {
                    'sheetId': report_sheet.id,
                    'startRowIndex': 1,
                    'endRowIndex': max_row,
                    'startColumnIndex': 0,
                    'endColumnIndex': 1
                },
                'cell': {
                    'userEnteredFormat': self._formats["left"]
                },
                'fields': 'userEnteredFormat'
            }
        })

        # Add request to format body currency cells
        body['requests'].append({
            'repeatCell': {
                'range': {
                    'sheetId': report_sheet.id,
                    'startRowIndex': 1,
                    'endRowIndex': max_row,
                    'startColumnIndex': 1,
                    'endColumnIndex': max_col
                },
                'cell': {
                    'userEnteredFormat': self._formats["center"]
                },
                'fields': 'userEnteredFormat'
            }
        })

        res = workbook.batch_update(body)

        return res
