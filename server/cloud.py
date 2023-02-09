import gspread
from google.oauth2.service_account import Credentials
from gspread_formatting import *
from datetime import date
import server.utils as utils
import logging


class CloudServices:

    ZONES = \
        [
            'Norte',
            'Sur',
            'Oriente',
            'Occidente',
            'Regiones'
        ]

    COLUMNS_SHEETS = \
        {
            'farmacia': 'Cliente',
            'factura': 'Factura',
            'cajas': 'Cajas',
            'bolsas': 'Bolsas',
            'lios': 'Lios',
            'valor': 'Valor a cobrar',
            'cobrado': 'Valor cobrado',
            'empacador': 'Empacador',
            'fecha': 'Hora empaque',
            'mensajero': 'Mensajero',
            'hora': 'Hora salida',
            'notas': 'Observaciones',
            'uid': 'uid'
        }

    MAX_COLS = len(COLUMNS_SHEETS)
    MAX_ROWS = 200

    def __init__(
            self,
            folder_id: str,
            json_path: str) -> None:
        """
        Create Google auth and Google Sheets objects
        """
        scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]

        credentials = Credentials.from_service_account_file(
            json_path,
            scopes=scopes
        )

        # Initialize sheets list
        self.__sheets = list()

        # Create client object
        self.client = gspread.authorize(credentials)

        # Initialize cloud services
        self._folder_id = folder_id

        self.__set_names()
        self.__initialize()

    @property
    def sheets(self) -> list:
        return self.__sheets

    @property
    def folder_id(self) -> list:
        return self._folder_id

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
            bg_rgb: tuple,
            fg_rgb: tuple,
            bold: bool,
            h_align: str = 'CENTER',
            v_align: str = 'MIDDLE',
            number_format: NumberFormat = None) -> dict:
        """
        Create cell format

        Args:
            bg_rgb (tuple): RGC background color
            fg_rgb (tuple): RGC foreground color
            bold (bool): true if bold
            h_align (str, optional): horizontal alignment. Defaults to 'CENTER'.
            v_align (str, optional): vertical alignment. Defaults to 'MIDDLE'.
            number_format (NumberFormat, optional): numerical format. Defaults to None.

        Returns:
            dict: specified CellFormat converted to dict
        """
        bg_clr = self.__get_color(*bg_rgb)
        fg_clr = self.__get_color(*fg_rgb)

        return CellFormat(
            backgroundColor=bg_clr,
            textFormat=TextFormat
            (
                bold=bold,
                foregroundColor=fg_clr,
            ),
            horizontalAlignment=h_align,
            verticalAlignment=v_align,
            numberFormat=number_format
        ).to_props()

    def __set_formats_specs(self) -> None:
        """
        Set cells formats
        """
        self._fmt_header = self.__create_format_cell(
            bg_rgb=(0, 150, 214),
            fg_rgb=(250, 250, 250),
            bold=True
        )

        self._fmt_body = self.__create_format_cell(
            bg_rgb=(255, 255, 255),
            fg_rgb=(69, 69, 69),
            bold=False
        )

        self._fmt_body_left = self.__create_format_cell(
            bg_rgb=(255, 255, 255),
            fg_rgb=(69, 69, 69),
            bold=False,
            h_align='LEFT'
        )

        self._fmt_body_currency = self.__create_format_cell(
            bg_rgb=(255, 255, 255),
            fg_rgb=(69, 69, 69),
            bold=False,
            number_format=NumberFormat(
                type='CURRENCY',
                pattern='$ #,###'
            )
        )

    def __set_names(self) -> None:
        """
        Set list with names of "Mensajeros" and "Empacadores"
        """
        try:
            spreadsheet = self.client.open_by_key(
                '1UwpyZ0_D_hp2pKU6uef4RI09MKZjUkcQXA9AsFnl3bw'
            )

            couriers = spreadsheet.get_worksheet(0)
            packers = spreadsheet.get_worksheet(1)

            data_couriers = couriers.get_all_records(head=1)
            data_packers = packers.get_all_records(head=1)

            self._names_couriers = [
                list(data.values())[0] for data in data_couriers
            ]

            self._names_packers = [
                list(data.values())[0] for data in data_packers
            ]

        except:
            logging.error(
                "Error al obtener nombres de mensajeros y empacadores",
                exc_info=True
            )

    def __initialize(self) -> None:
        """
        Create Workbook and Worksheets objects
        """
        filename = date.today().strftime("%Y-%m-%d")

        # Create CellFormat objects
        self.__set_formats_specs()

        # Create Google Sheets workbook if not exist
        if self.exist_file(filename):
            self.workbook = self.select_file(filename)
        else:
            self.workbook = self.client.create(
                title=filename,
                folder_id=self._folder_id
            )

            # Create all sheets
            self.__create_all_sheets()

    def __create_sheet(self, zone: str) -> gspread.Worksheet:
        """
        Args:
            zone (str): order region

        Returns:
            Worksheet: created gspread Worksheet
        """
        sheet = self.workbook.add_worksheet(
            title=zone,
            rows=CloudServices.MAX_ROWS,
            cols=CloudServices.MAX_COLS
        )

        sheet.clear()

        return sheet

    def __create_all_sheets(self) -> None:
        """
        Create five region sheets
        """
        try:
            self.__sheets = [
                self.__create_sheet(zone)
                for zone in CloudServices.ZONES
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

    def __set_columns_width(
            self,
            sheet: gspread.Worksheet) -> dict:
        """
        Set columns width

        Args:
            sheet (gspread.Worksheet): sheet object

        Returns:
            dict: Response batch_update method
        """
        body = {
            'requests': [
                {
                    'updateDimensionProperties': {
                        'range': {
                            'sheetId': sheet.id,
                            'dimension': 'COLUMNS',
                            'startIndex': 0,
                            'endIndex': 1
                        },
                        'properties': {
                            'pixelSize': 350
                        },
                        'fields': 'pixelSize'
                    }
                },
                {
                    'updateDimensionProperties': {
                        'range': {
                            'sheetId': sheet.id,
                            'dimension': 'COLUMNS',
                            'startIndex': 1,
                            'endIndex': 2
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
                            'sheetId': sheet.id,
                            'dimension': 'COLUMNS',
                            'startIndex': 2,
                            'endIndex': 5
                        },
                        'properties': {
                            'pixelSize': 70
                        },
                        'fields': 'pixelSize'
                    }
                },
                {
                    'updateDimensionProperties': {
                        'range': {
                            'sheetId': sheet.id,
                            'dimension': 'COLUMNS',
                            'startIndex': 5,
                            'endIndex': 7
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
                            'sheetId': sheet.id,
                            'dimension': 'COLUMNS',
                            'startIndex': 7,
                            'endIndex': 8
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
                            'sheetId': sheet.id,
                            'dimension': 'COLUMNS',
                            'startIndex': 8,
                            'endIndex': 9
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
                            'sheetId': sheet.id,
                            'dimension': 'COLUMNS',
                            'startIndex': 9,
                            'endIndex': 10
                        },
                        'properties': {
                            'pixelSize': 220
                        },
                        'fields': 'pixelSize'
                    }
                },
                {
                    'updateDimensionProperties': {
                        'range': {
                            'sheetId': sheet.id,
                            'dimension': 'COLUMNS',
                            'startIndex': 10,
                            'endIndex': 11
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
                            'sheetId': sheet.id,
                            'dimension': 'COLUMNS',
                            'startIndex': 11,
                            'endIndex': 12
                        },
                        'properties': {
                            'pixelSize': 350
                        },
                        'fields': 'pixelSize'
                    }
                },
                {
                    'updateDimensionProperties': {
                        'range': {
                            'sheetId': sheet.id,
                            'dimension': 'ROWS',
                            'startIndex': 0,
                            'endIndex': CloudServices.MAX_ROWS
                        },
                        'properties': {
                            'pixelSize': 28
                        },
                        'fields': 'pixelSize'
                    }
                },
                {
                    'updateDimensionProperties': {
                        'range': {
                            'sheetId': sheet.id,
                            'dimension': 'COLUMNS',
                            'startIndex': 12,
                            'endIndex': 13
                        },
                        "properties": {
                            "hiddenByUser": True,
                        },
                        "fields": 'hiddenByUser'
                    }
                },
                {
                    "setDataValidation": {
                        "range": {
                            "sheetId": sheet.id,
                            "startRowIndex": 1,
                            "endRowIndex": CloudServices.MAX_ROWS,
                            "startColumnIndex": 9,
                            "endColumnIndex": 10
                        },
                        "rule": {
                            "condition": {
                                "type": 'ONE_OF_LIST',
                                "values": [
                                    {"userEnteredValue": name} for name in self._names_couriers
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
                            "endRowIndex": CloudServices.MAX_ROWS,
                            "startColumnIndex": 7,
                            "endColumnIndex": 8
                        },
                        "rule": {
                            "condition": {
                                "type": 'ONE_OF_LIST',
                                "values": [
                                    {"userEnteredValue": name} for name in self._names_packers
                                ],
                            },
                            "showCustomUi": True,
                            "strict": True
                        }
                    }
                }
            ]
        }

        res = self.workbook.batch_update(body)

        return res

    def __set_cell_formats(
            self,
            sheet: gspread.Worksheet) -> dict:
        """
        Set rows height and format cells

        Args:
            sheet (gspread.Worksheet): sheet object

        Returns:
            dict: Response batch_update method
        """
        max_row = CloudServices.MAX_ROWS
        max_col = CloudServices.MAX_COLS

        body = {
            'requests': [
                {
                    'repeatCell': {
                        'range': {
                            'sheetId': sheet.id,
                            'startRowIndex': 0,
                            'endRowIndex': 1,
                            'startColumnIndex': 0,
                            'endColumnIndex': max_col
                        },
                        'cell': {
                            'userEnteredFormat': self._fmt_header
                        },
                        'fields': 'userEnteredFormat'
                    }
                },
                {
                    'repeatCell': {
                        'range': {
                            'sheetId': sheet.id,
                            'startRowIndex': 1,
                            'endRowIndex': max_row,
                            'startColumnIndex': 0,
                            'endColumnIndex': 1
                        },
                        'cell': {
                            'userEnteredFormat': self._fmt_body_left
                        },
                        'fields': 'userEnteredFormat'
                    }
                },
                {
                    'repeatCell': {
                        'range': {
                            'sheetId': sheet.id,
                            'startRowIndex': 1,
                            'endRowIndex': max_row,
                            'startColumnIndex': 1,
                            'endColumnIndex': 5
                        },
                        'cell': {
                            'userEnteredFormat': self._fmt_body
                        },
                        'fields': 'userEnteredFormat'
                    }
                },
                {
                    'repeatCell': {
                        'range': {
                            'sheetId': sheet.id,
                            'startRowIndex': 1,
                            'endRowIndex': max_row,
                            'startColumnIndex': 5,
                            'endColumnIndex': 7
                        },
                        'cell': {
                            'userEnteredFormat': self._fmt_body_currency
                        },
                        'fields': 'userEnteredFormat'
                    }
                },
                {
                    'repeatCell': {
                        'range': {
                            'sheetId': sheet.id,
                            'startRowIndex': 1,
                            'endRowIndex': max_row,
                            'startColumnIndex': 7,
                            'endColumnIndex': max_col
                        },
                        'cell': {
                            'userEnteredFormat': self._fmt_body
                        },
                        'fields': 'userEnteredFormat'
                    }
                }
            ]
        }

        res = self.workbook.batch_update(body)

        return res

    @staticmethod
    def get_col_name(col_idx: int) -> str:
        return chr(ord('@') + col_idx)

    def __format_sheet(
            self,
            sheet: gspread.Worksheet) -> None:
        """
        Apply format to each sheet

        sheet (gspread.Worksheet): sheet to format
        """
        # insert titles
        self.insert_row(
            list(CloudServices.COLUMNS_SHEETS.values()),
            sheet.title,
        )

        self.__set_columns_width(
            sheet
        )

        self.__set_cell_formats(
            sheet
        )

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

    def insert_row(
            self,
            values: list,
            zone: str) -> None:
        """
        Insert element into last row

        Args:
            values (list): customer invoice info
            zone (str): zone name (also Sheet title)
        """
        try:
            sheet = self.workbook.worksheet(zone)

            curr_row = len(sheet.col_values(1)) + 1

            # update current non written row
            sheet.update(
                f'A{curr_row}:M{curr_row}',
                [values]
            )

            # include basic filter in the header
            sheet.set_basic_filter(
                f'A1:M{curr_row}'
            )

        except Exception:
            logging.error(
                "Error al insertar nueva fila al archivo de Google Sheets",
                exc_info=True
            )

    def __find_row_by_key(
            self,
            sheet: gspread.Worksheet,
            key: str) -> int:
        """
        Find row index by database uid

        Args:
            key (str): database uid

        Returns:
            int: row index
        """
        cell = sheet.find(
            key,
            in_column=13  # uid column
        )

        return cell.row if cell is not None else -1

    def update_row(
            self,
            zone: str,
            key: str,
            values: list) -> None:
        """
        Update row if there are any changes

        Args:
            zone (str): region in database
            key (str): database key
            values (list): 2D list with row values
        """
        sheet = self.workbook.worksheet(zone)

        row = self.__find_row_by_key(
            sheet,
            key
        )

        if row != -1:
            range_name = f'A{row}:M{row}'
            row_values = sheet.row_values(row)

            values[0][6] = row_values[6]
            values[0][9] = row_values[9]
            values[0][10] = row_values[10]
            values[0][11] = row_values[11]

            sheet.update(
                range_name,
                values
            )

    def delete_row(
            self,
            key: str) -> list:
        """
        Delete row of specified region sheet and key

        Args:
            key (str): uid database
        """
        for zone in CloudServices.ZONES:
            sheet = self.workbook.worksheet(zone)
            row = self.__find_row_by_key(
                sheet,
                key
            )
            if row != -1:
                break

        # get values previous to delete
        deleted_values = sheet.row_values(row)

        sheet.delete_rows(row)

        return deleted_values

    def __get_sheet_names(self) -> list:
        """
        Get names of all sheets of current Spreadsheet

        Returns:
            list: names of sheets
        """
        return [ws.title for ws in self.workbook.worksheets()]

    def __get_all_values(self) -> list:
        """
        Get values of all sheets in one list

        Returns:
            list: contains all rows of all sheets
        """
        return sum(
            [
                self.workbook.worksheet(zone).get_all_values()[1:]
                for zone in CloudServices.ZONES
            ],
            []
        )

    def __exist_sheet(
            self,
            sheet_name: str) -> bool:
        """
        Check if sheet exists

        Args:
            sheet_name (str): sheet name

        Returns:
            bool: true if sheet exists
        """
        worksheet_names = self.__get_sheet_names()

        return sheet_name in worksheet_names

    def exist_file(self, file: str) -> bool:
        """
        Checks if file exists in specified folder

        Args:
            file (str): File name

        Returns:
            bool: True if there is a file in the folder
        """
        dict_files = self.client.list_spreadsheet_files(
            folder_id=self._folder_id
        )

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
            ss = self.client.open(
                title=filename,
                folder_id=self._folder_id
            )

            return ss

        except Exception as e:
            logging.error(
                f'Error {e} opening specified filename',
                exc_info=True
            )

    def __create_report_sheet(
            self,
            filename: str,
            name: str,
            headers: list,
            reset: bool) -> tuple[gspread.Spreadsheet, gspread.Worksheet, bool]:
        """
        Create report sheet

        Args:
            filename (str): Name of file to get report.
            name (str): Sheet name.
            headers (List[str]): List of headers for the sheet.
            reset (bool): True if reset sheet (only headers retained).

        Returns:
            tuple: Contains the Spreadsheet object, the Worksheet object, and a bool flag indicating if the sheet existed.
        """
        try:
            sheet_exists = self.__exist_sheet(name)
            workbook = self.select_file(filename)

            if sheet_exists:
                report_sheet = workbook.worksheet(name)

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

            return workbook, report_sheet, sheet_exists

        except Exception as e:
            logging.error(
                f'Error opening specified filename: {e}',
                exc_info=True
            )

            return None, None, False

    def _create_common_report(
            self,
            name: str,
            filename: str,
            email: str) -> dict:
        """
        Common format to create Dispatches or Treasury format

        Args:
            name (str): Report name ('treasury' or 'dispatches')
            filename (str): Name of the file to get the report
            email (str): Email account for whom the sheet is non-protected

        Returns:
            dict: Contains batch update response
        """
        headers = \
            [
                'Mensajero',
                'Zona',
                'Factura',
                'Cliente',
                'Valor cobrado',
                'Observaciones',
                'Notas Tesorería',
                'uid'
            ]

        workbook, report_sheet, sheet_exists = \
            self.__create_report_sheet(
                filename=filename,
                name=name,
                headers=headers,
                reset=True
            )

        if report_sheet is None:
            return {}

        code_values = report_sheet.col_values(2)
        start_row = len(code_values)

        values_to_append = []

        for zone in CloudServices.ZONES:
            sheet = workbook.worksheet(zone)
            rows = sheet.get_all_values()

            for row in rows[1:]:
                code = row[1]

                if code == 0 or code not in code_values[1:]:
                    if row[9] != '':
                        values = \
                            [
                                row[9],
                                sheet.title,
                                row[1],
                                row[0],
                                utils.str_to_int(row[6]),
                                row[11],
                                '',
                                row[12],
                            ]

                        values_to_append.append(values)

        if not values_to_append:
            return {}

        report_sheet.append_rows(
            values=values_to_append
        )

        max_row = len(report_sheet.col_values(1))
        max_col = len(headers)

        body = {'requests': []}

        if not sheet_exists:
            body['requests'].append({
                "addProtectedRange": {
                    "protectedRange": {
                        "range": {
                            "sheetId": report_sheet.id,
                        },
                        "editors": {
                            "domainUsersCanEdit": False,
                            "users": [email]
                        },
                        "warningOnly": False
                    }
                }
            })

            # hide uid column
            body['requests'].append({
                'updateDimensionProperties': {
                    'range': {
                        'sheetId': report_sheet.id,
                        'dimension': 'COLUMNS',
                        'startIndex': 6,
                        'endIndex': max_col
                    },
                    "properties": {
                        "hiddenByUser": True,
                    },
                    "fields": 'hiddenByUser'
                }
            })

            # add requests for updating column sizes
            body['requests'].append({
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
            })

            body['requests'].append({
                'updateDimensionProperties': {
                    'range': {
                        'sheetId': report_sheet.id,
                        'dimension': 'COLUMNS',
                        'startIndex': 1,
                        'endIndex': 3
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
                        'dimension': 'COLUMNS',
                        'startIndex': 3,
                        'endIndex': 4
                    },
                    'properties': {
                        'pixelSize': 350
                    },
                    'fields': 'pixelSize'
                }
            })

            body['requests'].append({
                'updateDimensionProperties': {
                    'range': {
                        'sheetId': report_sheet.id,
                        'dimension': 'COLUMNS',
                        'startIndex': 4,
                        'endIndex': 5
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
                        'startIndex': 5,
                        'endIndex': 6
                    },
                    'properties': {
                        'pixelSize': 450
                    },
                    'fields': 'pixelSize'
                }
            })

        # add request for updating row sizes
        body['requests'].append({
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
                    'userEnteredFormat': self._fmt_header
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
                    'userEnteredFormat': self._fmt_body_left
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
                    'startColumnIndex': 1,
                    'endColumnIndex': 3
                },
                'cell': {
                    'userEnteredFormat': self._fmt_body
                },
                'fields': 'userEnteredFormat'
            }
        })

        body['requests'].append({
            'repeatCell': {
                'range': {
                    'sheetId': report_sheet.id,
                    'startRowIndex': 1,
                    'endRowIndex': max_row,
                    'startColumnIndex': 3,
                    'endColumnIndex': 4
                },
                'cell': {
                    'userEnteredFormat': self._fmt_body_left
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
                    'startColumnIndex': 4,
                    'endColumnIndex': 5
                },
                'cell': {
                    'userEnteredFormat': self._fmt_body_currency
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
                    'startColumnIndex': 5,
                    'endColumnIndex': max_col
                },
                'cell': {
                    'userEnteredFormat': self._fmt_body
                },
                'fields': 'userEnteredFormat'
            }
        })

        res = workbook.batch_update(body)

        return res

    def create_treasury_report(self, filename: str) -> dict:
        """
        Args:
            filename (str): File name (date based)

        Returns:
            dict: Batch update response
        """
        self._create_common_report(
            name='Tesorería',
            filename=filename,
            email='depositoemestesoreria1@gmail.com'
        )

    def create_dispatches_report(self, filename: str) -> dict:
        """
        Args:
            filename (str): File name (date based)

        Returns:
            dict: Batch update response
        """
        self._create_common_report(
            name='Despachos',
            filename=filename,
            email='depositoemesdespachos1@gmail.com'
        )

    def create_couriers_report(self, filename: str) -> dict:
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
        data = \
            {
                name:
                {
                    "value": 0,
                    "count": 0,
                    "cajas": 0,
                    "bolsas": 0,
                    "lios": 0
                } for name in self._names_couriers
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
            values=[[k]+list(v.values()) for k, v in data.items()]
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
                    'userEnteredFormat': self._fmt_header
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
                    'userEnteredFormat': self._fmt_body_left
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
                    'userEnteredFormat': self._fmt_body_currency
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
                    'userEnteredFormat': self._fmt_body
                },
                'fields': 'userEnteredFormat'
            }
        })

        res = workbook.batch_update(body)

        return res

    def create_packers_report(self, filename: str) -> dict:
        """
        Create packers report

        Args:
            filename (str): Name of the file to get the report

        Returns:
            dict: Contains batch update response
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

        # store data
        data = \
            {
                name:
                {
                    "count": 0,
                    "cajas": 0,
                    "bolsas": 0,
                    "lios": 0
                } for name in self._names_packers + ['Otro']
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
            values=[[k]+list(v.values()) for k, v in data.items()]
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
                    'userEnteredFormat': self._fmt_header
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
                    'userEnteredFormat': self._fmt_body_left
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
                    'endColumnIndex': max_col
                },
                'cell': {
                    'userEnteredFormat': self._fmt_body
                },
                'fields': 'userEnteredFormat'
            }
        })

        res = workbook.batch_update(body)

        return res
