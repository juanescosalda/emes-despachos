import os
from server.dispatches import EmesDispatch
from googleapiclient.discovery import build
import gspread
from google.oauth2.service_account import Credentials
from gspread_formatting import *
from apiclient import errors
import server.utils as utils
import locale

locale.setlocale(locale.LC_ALL, 'en_US.UTF8')

ZONES = [
    'Norte',
    'Sur',
    'Oriente',
    'Occidente',
    'Regiones'
]

name_couriers = [
    'Elmer',
    'Leon Jaime Muriel',
    'Andres Villada',
    'Juan David Cataño',
    'Sebastian Chica Araque',
    'Sergio Gutierrez',
    'Sergio Muñoz',
    'Ivan Ospina',
    'Edison Ruiz',
    'Jhon Rojas'
]

name_packers = [
    'Aldair Quintana',
    'Anderson Castañeda',
    'Camila Borja',
    'Cesar Palacio',
    'David Rojas',
    'Diego Moncada',
    'Felipe Mejía',
    'Jhoan Felipe Lazaro',
    'Jhon Ever Ortiz',
    'Juan Diego Rodas',
    'Juan Pablo Flores',
    'Julian Cardona',
    'Julian Castro',
    'Sebastián Flores'
]

json_path_gs = fr'C:\Users\{os.getlogin()}\Desktop\Emes despachos\emes-empaques-gspread.json'

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

credentials = Credentials.from_service_account_file(
    json_path_gs,
    scopes=scopes
)

# create client object
client = gspread.authorize(credentials)

folder_id = '1S4mTxoIFIKNibVKRvofJCyLTGeOW0ye4'

global workbook

workbook = client.open_by_key(
    '1vtPP26Rgh9UPpGS2hcGspEPRrZmi2CwkJtyxdMAcDAY'
)

worksheet_names = [ws.title for ws in workbook.worksheets()]


def get_color(red: int, blue: int, green: int):
    return Color(
        utils.rgb2p(red),
        utils.rgb2p(blue),
        utils.rgb2p(green)
    )


def create_format_cell(
        bg_rgb: tuple,
        fg_rgb: tuple,
        bold: bool,
        h_align: str = 'CENTER',
        v_align: str = 'MIDDLE',
        number_format: NumberFormat = None) -> dict:
    bg_clr = get_color(*bg_rgb)
    fg_clr = get_color(*fg_rgb)

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


fmt_header = create_format_cell(
    bg_rgb=(0, 150, 214),
    fg_rgb=(250, 250, 250),
    bold=True
)

fmt_body = create_format_cell(
    bg_rgb=(255, 255, 255),
    fg_rgb=(69, 69, 69),
    bold=False
)

fmt_body_left = create_format_cell(
    bg_rgb=(255, 255, 255),
    fg_rgb=(69, 69, 69),
    bold=False,
    h_align='LEFT',
)

fmt_body_currency = create_format_cell(
    bg_rgb=(255, 255, 255),
    fg_rgb=(69, 69, 69),
    bold=False,
    number_format=NumberFormat(
        type='CURRENCY',
        pattern='$ #,###'
    )
)


def main():
    create_dispatches_report('2023-02-02')
    # db_url = 'https://emes-empaques-default-rtdb.firebaseio.com/'
    # folder_id_planillas = '1S4mTxoIFIKNibVKRvofJCyLTGeOW0ye4'
    # folder_id_remisiones = '1A2UP-JKrQvJV0SCMSD0IDa3ts-uOUJVR'
    # json_path_db = fr'C:\Users\{os.getlogin()}\Desktop\Emes despachos\emes-empaques-718c8b73c894.json'
    # json_path_gs = fr'C:\Users\{os.getlogin()}\Desktop\Emes despachos\emes-empaques-gspread.json'

    # emes = EmesDispatch(
    #     db_url,
    #     folder_id_planillas,
    #     folder_id_remisiones,
    #     json_path_db,
    #     json_path_gs
    # )
    # emes.connect()


def exist_sheet(sheet_name: str):
    return sheet_name in worksheet_names


def create_report_sheet(filename: str, name: str, headers: list, reset: bool):
    sheet_exists = exist_sheet(name)

    workbook = select_file(filename)

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

        # insert headers
        report_sheet.append_row(
            values=headers
        )

    return report_sheet, sheet_exists


def exist_file(file: str):
    dict_files = client.list_spreadsheet_files(
        folder_id=folder_id
    )

    files = [d['name'] for d in dict_files]

    return file in files


def select_file(filename: str):
    if exist_file(filename):
        return client.open(
            title=filename,
            folder_id=folder_id
        )
    else:
        return None  # revisar qué pasa si no existe


def get_client_name(client: str):
    separator = '-'

    if client.find(separator) != -1:
        parts = client.split(separator)
        return parts[0].strip()
    else:
        return client


def create_dispatches_report(filename: str) -> dict:
    """
    Create treasury report

    Returns:
        dict: contains batch update response
    """
    headers = \
        [
            'Mensajero',
            'Zona',
            'Factura',
            'Cliente',
            'Valor cobrado',
            'Observaciones',
            'Aprobado',
            'Notas Tesorería',
            'uid'
        ]

    report_sheet, sheet_exists = create_report_sheet(
        filename=filename,
        name='Despachos',
        headers=headers,
        reset=True
    )

    code_values = report_sheet.col_values(2)
    start_row = len(code_values)

    values_to_append = []

    for zone in ZONES:
        sheet = workbook.worksheet(zone)
        rows = sheet.get_all_values()

        for row in rows[1:]:
            code = row[1]

            if code == 0 or code not in code_values[1:]:
                client = get_client_name(row[0])
                if row[9] != '':
                    values = \
                        [
                            row[9],
                            sheet.title,
                            row[1],
                            client,
                            utils.str_to_int(row[6]),
                            row[11],
                            False,
                            '',
                            row[12],
                        ]

                    values_to_append.append(values)

    if not values_to_append:
        return

    report_sheet.append_rows(
        values=values_to_append
    )

    max_row = len(report_sheet.col_values(1))
    max_col = len(headers)

    body = {'requests': []}

    if not sheet_exists:
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
                    'pixelSize': 300
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

    # NUEVO CHECKLIST
    body['requests'].append({
        # "setDataValidation": {
        #     'range': {
        #         'sheetId': sheet.id,
        #         'startRowIndex': 1,
        #         'endRowIndex': max_row,
        #         'startColumnIndex': 5,
        #         'endColumnIndex': 6
        #     },
        #     'rule': {
        #         'condition': {
        #             'type': 'BOOLEAN',
        #             'values': [],
        #         },
        #         'showCustomUi': True,
        #         'strict': False
        #     }
        # },
        # "repeatCell": {
        #     'range': {
        #         'sheetId': sheet.id,
        #         'startRowIndex': 1,
        #         'endRowIndex': max_row,
        #         'startColumnIndex': 5,
        #         'endColumnIndex': 6
        #     },
        #     "cell": {
        #         "userEnteredValue": {
        #             "boolValue": False
        #         }
        #     },
        #     "fields": "userEnteredValue"
        # }
        # "updateCells": {
        #     "rows": [{"values": [{"userEnteredValue": {"formulaValue": "=IF(A1, \"✔\", \"\")"}}]}],
        #     "fields": "userEnteredValue",
        #     'range': {
        #         "sheetId": sheet.id,
        #         "startRowIndex": 1,
        #         "endRowIndex": max_row,
        #         "startColumnIndex": 5,
        #         "endColumnIndex": 6
        #     }
        # }
        "updateCells": {
            "rows": [
                {
                    "values":
                    [
                        {
                            "userEnteredFormat":
                            {
                                "checkbox": True
                            }
                        }
                        for _ in range(1, max_row)
                    ]
                }
            ],
            "fields": "userEnteredValue.checkboxValue",
            'range': {
                "sheetId": sheet.id,
                "startRowIndex": 1,
                "endRowIndex": max_row,
                "startColumnIndex": 5,
                "endColumnIndex": 6
            }
        }
        # 'repeatCell': {
        #     'cell': {'dataValidation': {'condition': {'type': 'BOOLEAN'}}},
        #     'range': {
        #         "sheetId": sheet.id,
        #         "startRowIndex": 1,
        #         "endRowIndex": max_row,
        #         "startColumnIndex": 5,
        #         "endColumnIndex": 6
        #     },
        #     'fields': 'dataValidation'
        # }
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
                'userEnteredFormat': fmt_header
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
                'userEnteredFormat': fmt_body_left
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
                'userEnteredFormat': fmt_body
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
                'userEnteredFormat': fmt_body_left
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
                'userEnteredFormat': fmt_body_currency
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
                'userEnteredFormat': fmt_body
            },
            'fields': 'userEnteredFormat'
        }
    })

    res = workbook.batch_update(body)

    return res

# def create_treasury_report() -> dict:
#     """
#     Create treasury report

#     Returns:
#         dict: contains batch update response
#     """
#     headers = \
#         [
#             'Mensajero',
#             'Factura',
#             'Cliente',
#             'Valor cobrado',
#             'Observaciones',
#             'Notas Tesorería',
#             'uid'
#         ]

#     report_sheet, sheet_exists = create_report_sheet(
#         name='Tesorería',
#         headers=headers,
#         reset=False
#     )

#     if sheet_exists:
#         color_theme = \
#             {
#                 "red": utils.rgb2p(216),
#                 "green": utils.rgb2p(255),
#                 "blue": utils.rgb2p(223)
#             }
#     else:
#         color_theme = \
#             {
#                 "red": utils.rgb2p(216),
#                 "green": utils.rgb2p(223),
#                 "blue": utils.rgb2p(255)
#             }

#     code_values = report_sheet.col_values(2)
#     start_row = len(code_values)

#     values_to_append = []

#     for zone in ZONES:
#         sheet = workbook.worksheet(zone)
#         rows = sheet.get_all_values()

#         for row in rows[1:]:
#             code = row[1]

#             if code == 0 or code not in code_values[1:]:
#                 if row[9] != '':
#                     values = \
#                         [
#                             row[9],
#                             row[1],
#                             row[0],
#                             utils.str_to_int(row[6]),
#                             row[11],
#                             '',
#                             row[12],
#                         ]

#                     values_to_append.append(values)

#     if not values_to_append:
#         return

#     report_sheet.append_rows(
#         values=values_to_append
#     )

#     max_row = len(report_sheet.col_values(1))
#     max_col = len(headers)

#     body = {'requests': []}

#     if not sheet_exists:
#         body['requests'].append({
#             "addProtectedRange": {
#                 "protectedRange": {
#                     "range": {
#                         "sheetId": report_sheet.id,
#                     },
#                     "editors": {
#                         "domainUsersCanEdit": False,
#                         "users": [
#                             'depositoemestesoreria1@gmail.com',
#                         ]
#                     },
#                     "warningOnly": False
#                 }
#             }
#         })

#         # hide uid column
#         body['requests'].append({
#             'updateDimensionProperties': {
#                 'range': {
#                     'sheetId': report_sheet.id,
#                     'dimension': 'COLUMNS',
#                     'startIndex': 6,
#                     'endIndex': max_col
#                 },
#                 "properties": {
#                     "hiddenByUser": True,
#                 },
#                 "fields": 'hiddenByUser'
#             }
#         })

#         # add requests for updating column sizes
#         body['requests'].append({
#             'updateDimensionProperties': {
#                 'range': {
#                     'sheetId': report_sheet.id,
#                     'dimension': 'COLUMNS',
#                     'startIndex': 0,
#                     'endIndex': 1
#                 },
#                 'properties': {
#                     'pixelSize': 180
#                 },
#                 'fields': 'pixelSize'
#             }
#         })

#         body['requests'].append({
#             'updateDimensionProperties': {
#                 'range': {
#                     'sheetId': report_sheet.id,
#                     'dimension': 'COLUMNS',
#                     'startIndex': 1,
#                     'endIndex': 2
#                 },
#                 'properties': {
#                     'pixelSize': 100
#                 },
#                 'fields': 'pixelSize'
#             }
#         })

#         body['requests'].append({
#             'updateDimensionProperties': {
#                 'range': {
#                     'sheetId': report_sheet.id,
#                     'dimension': 'COLUMNS',
#                     'startIndex': 2,
#                     'endIndex': 3
#                 },
#                 'properties': {
#                     'pixelSize': 350
#                 },
#                 'fields': 'pixelSize'
#             }
#         })

#         body['requests'].append({
#             'updateDimensionProperties': {
#                 'range': {
#                     'sheetId': report_sheet.id,
#                     'dimension': 'COLUMNS',
#                     'startIndex': 3,
#                     'endIndex': 4
#                 },
#                 'properties': {
#                     'pixelSize': 140
#                 },
#                 'fields': 'pixelSize'
#             }
#         })

#         body['requests'].append({
#             'updateDimensionProperties': {
#                 'range': {
#                     'sheetId': report_sheet.id,
#                     'dimension': 'COLUMNS',
#                     'startIndex': 4,
#                     'endIndex': 5
#                 },
#                 'properties': {
#                     'pixelSize': 300
#                 },
#                 'fields': 'pixelSize'
#             }
#         })

#         body['requests'].append({
#             'updateDimensionProperties': {
#                 'range': {
#                     'sheetId': report_sheet.id,
#                     'dimension': 'COLUMNS',
#                     'startIndex': 5,
#                     'endIndex': 6
#                 },
#                 'properties': {
#                     'pixelSize': 300
#                 },
#                 'fields': 'pixelSize'
#             }
#         })

#     # add request for updating row sizes
#     body['requests'].append({
#         'updateDimensionProperties': {
#             'range': {
#                 'sheetId': report_sheet.id,
#                 'dimension': 'ROWS',
#                 'startIndex': start_row - 1,
#                 'endIndex': max_row
#             },
#             'properties': {
#                 'pixelSize': 30
#             },
#             'fields': 'pixelSize'
#         }
#     })

#     # add basic filter
#     body['requests'].append({
#         'setBasicFilter': {
#             'filter': {
#                 'range': {
#                     'sheetId': report_sheet.id,
#                     'startRowIndex': 0,
#                     'endRowIndex': max_row,
#                     'startColumnIndex': 0,
#                     'endColumnIndex': max_col
#                 },
#             }
#         }
#     })

#     # add request to format header cells
#     body['requests'].append({
#         'repeatCell': {
#             'range': {
#                 'sheetId': report_sheet.id,
#                 'startRowIndex': 0,
#                 'endRowIndex': 1,
#                 'startColumnIndex': 0,
#                 'endColumnIndex': max_col
#             },
#             'cell': {
#                 'userEnteredFormat': fmt_header
#             },
#             'fields': 'userEnteredFormat'
#         }
#     })

#     # add request to format body left cells
#     body['requests'].append({
#         'repeatCell': {
#             'range': {
#                 'sheetId': report_sheet.id,
#                 'startRowIndex': 1,
#                 'endRowIndex': max_row,
#                 'startColumnIndex': 0,
#                 'endColumnIndex': 1
#             },
#             'cell': {
#                 'userEnteredFormat': fmt_body_left
#             },
#             'fields': 'userEnteredFormat'
#         }
#     })

#     # add request to format body cells
#     body['requests'].append({
#         'repeatCell': {
#             'range': {
#                 'sheetId': report_sheet.id,
#                 'startRowIndex': 1,
#                 'endRowIndex': max_row,
#                 'startColumnIndex': 1,
#                 'endColumnIndex': 2
#             },
#             'cell': {
#                 'userEnteredFormat': fmt_body
#             },
#             'fields': 'userEnteredFormat'
#         }
#     })

#     body['requests'].append({
#         'repeatCell': {
#             'range': {
#                 'sheetId': report_sheet.id,
#                 'startRowIndex': 1,
#                 'endRowIndex': max_row,
#                 'startColumnIndex': 2,
#                 'endColumnIndex': 3
#             },
#             'cell': {
#                 'userEnteredFormat': fmt_body_left
#             },
#             'fields': 'userEnteredFormat'
#         }
#     })

#     # add request to format body currency cells
#     body['requests'].append({
#         'repeatCell': {
#             'range': {
#                 'sheetId': report_sheet.id,
#                 'startRowIndex': 1,
#                 'endRowIndex': max_row,
#                 'startColumnIndex': 3,
#                 'endColumnIndex': 4
#             },
#             'cell': {
#                 'userEnteredFormat': fmt_body_currency
#             },
#             'fields': 'userEnteredFormat'
#         }
#     })

#     # add request to format body cells
#     body['requests'].append({
#         'repeatCell': {
#             'range': {
#                 'sheetId': report_sheet.id,
#                 'startRowIndex': 1,
#                 'endRowIndex': max_row,
#                 'startColumnIndex': 4,
#                 'endColumnIndex': max_col
#             },
#             'cell': {
#                 'userEnteredFormat': fmt_body
#             },
#             'fields': 'userEnteredFormat'
#         }
#     })

#     # add request to set background color
#     body['requests'].append({
#         'repeatCell': {
#             'range': {
#                 'sheetId': report_sheet.id,
#                 'startRowIndex': start_row,
#                 'endRowIndex': max_row,
#                 'startColumnIndex': 0,
#                 'endColumnIndex': max_col
#             },
#             'cell': {
#                 'userEnteredFormat': {
#                     'backgroundColor': color_theme
#                 }
#             },
#             'fields': 'userEnteredFormat.backgroundColor'
#         }
#     })

#     res = workbook.batch_update(body)

#     return res

# def create_packers_report():
#     report_sheet, sheet_exists = create_report_sheet(
#         name='Empacadores',
#         headers=[
#             'Empacador',
#             '# Empaques',
#             'Cajas',
#             'Bolsas',
#             'Lios'
#         ],
#         reset=True
#     )

#     # store data
#     data = \
#         {
#             name:
#             {
#                 "count": 0,
#                 "cajas": 0,
#                 "bolsas": 0,
#                 "lios": 0
#             } for name in name_packers
#         }

#     rows = sum(
#         [
#             workbook.worksheet(zone).get_all_values()[1:]
#             for zone in ZONES
#         ],
#         []
#     )

#     for row in rows:
#         name = row[7]
#         if name != '':
#             data[name]['count'] += 1
#             data[name]['cajas'] += int(row[2])
#             data[name]['bolsas'] += int(row[3])
#             data[name]['lios'] += int(row[4])

#     report_sheet.append_rows(
#         values=[[k]+list(v.values()) for k, v in data.items()]
#     )

#     # get dimensiones
#     max_row = len(report_sheet.col_values(1))
#     max_col = len(data.keys()) + 1

#     # create requests
#     body = {'requests': []}

#     # protect sheet
#     if not sheet_exists:
#         body['requests'].append({
#             "addProtectedRange": {
#                 "protectedRange": {
#                     "range": {
#                         "sheetId": report_sheet.id,
#                     },
#                     "editors": {
#                         "domainUsersCanEdit": False,
#                         "users": [
#                             'depositoemesdespachos1@gmail.com',
#                             'jigomez6025@gmail.com'
#                         ]
#                     },
#                     "warningOnly": False
#                 }
#             }
#         })

#     body['requests'].append({
#         'updateDimensionProperties': {
#             'range': {
#                 'sheetId': report_sheet.id,
#                 'dimension': 'COLUMNS',
#                 'startIndex': 0,
#                 'endIndex': 1
#             },
#             'properties': {
#                 'pixelSize': 250
#             },
#             'fields': 'pixelSize'
#         }
#     })

#     body['requests'].append({
#         'updateDimensionProperties': {
#             'range': {
#                 'sheetId': report_sheet.id,
#                 'dimension': 'COLUMNS',
#                 'startIndex': 1,
#                 'endIndex': 2
#             },
#             'properties': {
#                 'pixelSize': 140
#             },
#             'fields': 'pixelSize'
#         }
#     })

#     body['requests'].append({
#         'updateDimensionProperties': {
#             'range': {
#                 'sheetId': report_sheet.id,
#                 'dimension': 'COLUMNS',
#                 'startIndex': 2,
#                 'endIndex': max_col
#             },
#             'properties': {
#                 'pixelSize': 100
#             },
#             'fields': 'pixelSize'
#         }
#     })

#     body['requests'].append({
#         'updateDimensionProperties': {
#             'range': {
#                 'sheetId': report_sheet.id,
#                 'dimension': 'ROWS',
#                 'startIndex': 0,
#                 'endIndex': max_row
#             },
#             'properties': {
#                 'pixelSize': 30
#             },
#             'fields': 'pixelSize'
#         }
#     })

#     body['requests'].append({
#         'setBasicFilter': {
#             'filter': {
#                 'range': {
#                     'sheetId': report_sheet.id,
#                     'startRowIndex': 0,
#                     'endRowIndex': max_row,
#                     'startColumnIndex': 0,
#                     'endColumnIndex': max_col
#                 },
#             }
#         }
#     })

#     # add request to format header cells
#     body['requests'].append({
#         'repeatCell': {
#             'range': {
#                 'sheetId': report_sheet.id,
#                 'startRowIndex': 0,
#                 'endRowIndex': 1,
#                 'startColumnIndex': 0,
#                 'endColumnIndex': max_col
#             },
#             'cell': {
#                 'userEnteredFormat': fmt_header
#             },
#             'fields': 'userEnteredFormat'
#         }
#     })

#     # add request to format body left cells
#     body['requests'].append({
#         'repeatCell': {
#             'range': {
#                 'sheetId': report_sheet.id,
#                 'startRowIndex': 1,
#                 'endRowIndex': max_row,
#                 'startColumnIndex': 0,
#                 'endColumnIndex': 1
#             },
#             'cell': {
#                 'userEnteredFormat': fmt_body_left
#             },
#             'fields': 'userEnteredFormat'
#         }
#     })

#     # add request to format body currency cells
#     body['requests'].append({
#         'repeatCell': {
#             'range': {
#                 'sheetId': report_sheet.id,
#                 'startRowIndex': 1,
#                 'endRowIndex': max_row,
#                 'startColumnIndex': 1,
#                 'endColumnIndex': max_col
#             },
#             'cell': {
#                 'userEnteredFormat': fmt_body
#             },
#             'fields': 'userEnteredFormat'
#         }
#     })

#     res = workbook.batch_update(body)

#     return res

# def create_couriers_report():
    # report_sheet, sheet_exists = create_report_sheet(
    #     name='Mensajeros',
    #     headers=[
    #         'Mensajero',
    #         'Valor cobrado',
    #         'Transportes',
    #         'Cajas',
    #         'Bolsas',
    #         'Lios'
    #     ],
    #     reset=True
    # )

    # # store data
    # data = \
    #     {
    #         name:
    #         {
    #             "value": 0,
    #             "count": 0,
    #             "cajas": 0,
    #             "bolsas": 0,
    #             "lios": 0
    #         } for name in name_couriers
    #     }

    # rows = sum(
    #     [
    #         workbook.worksheet(zone).get_all_values()[1:]
    #         for zone in ZONES
    #     ],
    #     []
    # )

    # # iterate over each row, and get the values
    # for row in rows:
    #     name = row[9]
    #     if name != '' and name != 'Mensajero':  # discard header REVISAR
    #         data[name]['value'] += utils.str_to_int(row[6])
    #         data[name]['count'] += 1
    #         data[name]['cajas'] += int(row[2])
    #         data[name]['bolsas'] += int(row[3])
    #         data[name]['lios'] += int(row[4])

    # # append multiple rows at once
    # report_sheet.append_rows(
    #     values=[[k]+list(v.values()) for k, v in data.items()]
    # )

    # # get dimensiones
    # max_row = len(report_sheet.col_values(1))
    # max_col = len(data.keys()) + 1

    # # create requests
    # body = {'requests': []}

    # # protect sheet
    # if not sheet_exists:
    #     body['requests'].append({
    #         "addProtectedRange": {
    #             "protectedRange": {
    #                 "range": {
    #                     "sheetId": report_sheet.id,
    #                 },
    #                 "editors": {
    #                     "domainUsersCanEdit": False,
    #                     "users": [
    #                         'depositoemesdespachos1@gmail.com',
    #                         'jigomez6025@gmail.com'
    #                     ]
    #                 },
    #                 "warningOnly": False
    #             }
    #         }
    #     })

    # body['requests'].append({
    #     'updateDimensionProperties': {
    #         'range': {
    #             'sheetId': report_sheet.id,
    #             'dimension': 'COLUMNS',
    #             'startIndex': 0,
    #             'endIndex': 1
    #         },
    #         'properties': {
    #             'pixelSize': 250
    #         },
    #         'fields': 'pixelSize'
    #     }
    # })

    # body['requests'].append({
    #     'updateDimensionProperties': {
    #         'range': {
    #             'sheetId': report_sheet.id,
    #             'dimension': 'COLUMNS',
    #             'startIndex': 1,
    #             'endIndex': 2
    #         },
    #         'properties': {
    #             'pixelSize': 140
    #         },
    #         'fields': 'pixelSize'
    #     }
    # })

    # body['requests'].append({
    #     'updateDimensionProperties': {
    #         'range': {
    #             'sheetId': report_sheet.id,
    #             'dimension': 'COLUMNS',
    #             'startIndex': 2,
    #             'endIndex': max_col
    #         },
    #         'properties': {
    #             'pixelSize': 100
    #         },
    #         'fields': 'pixelSize'
    #     }
    # })

    # body['requests'].append({
    #     'updateDimensionProperties': {
    #         'range': {
    #             'sheetId': report_sheet.id,
    #             'dimension': 'ROWS',
    #             'startIndex': 0,
    #             'endIndex': max_row
    #         },
    #         'properties': {
    #             'pixelSize': 30
    #         },
    #         'fields': 'pixelSize'
    #     }
    # })

    # # add basic filter
    # body['requests'].append({
    #     'setBasicFilter': {
    #         'filter': {
    #             'range': {
    #                 'sheetId': report_sheet.id,
    #                 'startRowIndex': 0,
    #                 'endRowIndex': max_row,
    #                 'startColumnIndex': 0,
    #                 'endColumnIndex': max_col
    #             },
    #         }
    #     }
    # })

    # # add request to format header cells
    # body['requests'].append({
    #     'repeatCell': {
    #         'range': {
    #             'sheetId': report_sheet.id,
    #             'startRowIndex': 0,
    #             'endRowIndex': 1,
    #             'startColumnIndex': 0,
    #             'endColumnIndex': max_col
    #         },
    #         'cell': {
    #             'userEnteredFormat': fmt_header
    #         },
    #         'fields': 'userEnteredFormat'
    #     }
    # })

    # # add request to format body left cells
    # body['requests'].append({
    #     'repeatCell': {
    #         'range': {
    #             'sheetId': report_sheet.id,
    #             'startRowIndex': 1,
    #             'endRowIndex': max_row,
    #             'startColumnIndex': 0,
    #             'endColumnIndex': 1
    #         },
    #         'cell': {
    #             'userEnteredFormat': fmt_body_left
    #         },
    #         'fields': 'userEnteredFormat'
    #     }
    # })

    # # add request to format body currency cells
    # body['requests'].append({
    #     'repeatCell': {
    #         'range': {
    #             'sheetId': report_sheet.id,
    #             'startRowIndex': 1,
    #             'endRowIndex': max_row,
    #             'startColumnIndex': 1,
    #             'endColumnIndex': 2
    #         },
    #         'cell': {
    #             'userEnteredFormat': fmt_body_currency
    #         },
    #         'fields': 'userEnteredFormat'
    #     }
    # })

    # # add request to format body cells
    # body['requests'].append({
    #     'repeatCell': {
    #         'range': {
    #             'sheetId': report_sheet.id,
    #             'startRowIndex': 1,
    #             'endRowIndex': max_row,
    #             'startColumnIndex': 2,
    #             'endColumnIndex': max_col
    #         },
    #         'cell': {
    #             'userEnteredFormat': fmt_body
    #         },
    #         'fields': 'userEnteredFormat'
    #     }
    # })

    # res = workbook.batch_update(body)

    # return res

# def create_treasury_report():
#     report_sheet, sheet_exists = create_report_sheet(
#         name='Tesorería',
#         headers=[
#             'Mensajero',
#             'Valor cobrado',
#             'Observaciones',
#             'uid'
#         ],
#         reset=False
#     )

#     if sheet_exists:
#         color_theme = \
#             {
#                 "red": utils.rgb2p(216),
#                 "green": utils.rgb2p(255),
#                 "blue": utils.rgb2p(223)
#             }
#     else:
#         color_theme = \
#             {
#                 "red": utils.rgb2p(216),
#                 "green": utils.rgb2p(223),
#                 "blue": utils.rgb2p(255)
#             }

#     uid_values = report_sheet.col_values(4)
#     start_row = len(uid_values)

#     values_to_append = []

#     for zone in ZONES:
#         sheet = workbook.worksheet(zone)
#         rows = sheet.get_all_values()

#         for row in rows[1:]:
#             if row[12] not in uid_values[1:]:
#                 if row[9] != '':
#                     values = \
#                         [
#                             row[9],
#                             utils.str_to_int(row[6]),
#                             row[11],
#                             row[12]
#                         ]

#                     values_to_append.append(values)

#     if not values_to_append:
#         return

#     report_sheet.append_rows(
#         values=values_to_append
#     )

#     max_row = len(report_sheet.col_values(1))
#     max_col = len(values_to_append[0])

#     body = {'requests': []}

#     if not sheet_exists:
#         body['requests'].append({
#             "addProtectedRange": {
#                 "protectedRange": {
#                     "range": {
#                         "sheetId": report_sheet.id,
#                     },
#                     "editors": {
#                         "domainUsersCanEdit": False,
#                         "users": [
#                             'depositoemesdespachos1@gmail.com',
#                             'depositoemestesoreria1@gmail.com',
#                             'depositoemescontabilidad1@gmail.com',
#                             'jigomez6025@gmail.com'
#                         ]
#                     },
#                     "warningOnly": False
#                 }
#             }
#         })

#         # hide uid column
#         body['requests'].append({
#             'updateDimensionProperties': {
#                 'range': {
#                     'sheetId': report_sheet.id,
#                     'dimension': 'COLUMNS',
#                     'startIndex': 3,
#                     'endIndex': 4
#                 },
#                 "properties": {
#                     "hiddenByUser": True,
#                 },
#                 "fields": 'hiddenByUser'
#             }
#         })

#         # add requests for updating column sizes
#         body['requests'].append({
#             'updateDimensionProperties': {
#                 'range': {
#                     'sheetId': report_sheet.id,
#                     'dimension': 'COLUMNS',
#                     'startIndex': 0,
#                     'endIndex': 1
#                 },
#                 'properties': {
#                     'pixelSize': 250
#                 },
#                 'fields': 'pixelSize'
#             }
#         })

#         body['requests'].append({
#             'updateDimensionProperties': {
#                 'range': {
#                     'sheetId': report_sheet.id,
#                     'dimension': 'COLUMNS',
#                     'startIndex': 1,
#                     'endIndex': 2
#                 },
#                 'properties': {
#                     'pixelSize': 140
#                 },
#                 'fields': 'pixelSize'
#             }
#         })

#         body['requests'].append({
#             'updateDimensionProperties': {
#                 'range': {
#                     'sheetId': report_sheet.id,
#                     'dimension': 'COLUMNS',
#                     'startIndex': 2,
#                     'endIndex': 3
#                 },
#                 'properties': {
#                     'pixelSize': 350
#                 },
#                 'fields': 'pixelSize'
#             }
#         })

#     # add request for updating row sizes
#     body['requests'].append({
#         'updateDimensionProperties': {
#             'range': {
#                 'sheetId': report_sheet.id,
#                 'dimension': 'ROWS',
#                 'startIndex': start_row - 1,
#                 'endIndex': max_row
#             },
#             'properties': {
#                 'pixelSize': 30
#             },
#             'fields': 'pixelSize'
#         }
#     })

#     # add basic filter
#     body['requests'].append({
#         'setBasicFilter': {
#             'filter': {
#                 'range': {
#                     'sheetId': report_sheet.id,
#                     'startRowIndex': 0,
#                     'endRowIndex': max_row,
#                     'startColumnIndex': 0,
#                     'endColumnIndex': max_col
#                 },
#             }
#         }
#     })

#     # add request to format header cells
#     body['requests'].append({
#         'repeatCell': {
#             'range': {
#                 'sheetId': report_sheet.id,
#                 'startRowIndex': 0,
#                 'endRowIndex': 1,
#                 'startColumnIndex': 0,
#                 'endColumnIndex': 3
#             },
#             'cell': {
#                 'userEnteredFormat': fmt_header.to_props()
#             },
#             'fields': 'userEnteredFormat'
#         }
#     })

#     # add request to format body left cells
#     body['requests'].append({
#         'repeatCell': {
#             'range': {
#                 'sheetId': report_sheet.id,
#                 'startRowIndex': 1,
#                 'endRowIndex': max_row,
#                 'startColumnIndex': 0,
#                 'endColumnIndex': 1
#             },
#             'cell': {
#                 'userEnteredFormat': fmt_body_left.to_props()
#             },
#             'fields': 'userEnteredFormat'
#         }
#     })

#     # add request to format body currency cells
#     body['requests'].append({
#         'repeatCell': {
#             'range': {
#                 'sheetId': report_sheet.id,
#                 'startRowIndex': 1,
#                 'endRowIndex': max_row,
#                 'startColumnIndex': 1,
#                 'endColumnIndex': 2
#             },
#             'cell': {
#                 'userEnteredFormat': fmt_body_currency.to_props()
#             },
#             'fields': 'userEnteredFormat'
#         }
#     })

#     # add request to format body cells
#     body['requests'].append({
#         'repeatCell': {
#             'range': {
#                 'sheetId': report_sheet.id,
#                 'startRowIndex': 1,
#                 'endRowIndex': max_row,
#                 'startColumnIndex': 2,
#                 'endColumnIndex': 3
#             },
#             'cell': {
#                 'userEnteredFormat': fmt_body.to_props()
#             },
#             'fields': 'userEnteredFormat'
#         }
#     })

#     # add request to set background color
#     body['requests'].append({
#         'repeatCell': {
#             'range': {
#                 'sheetId': report_sheet.id,
#                 'startRowIndex': start_row,
#                 'endRowIndex': max_row,
#                 'startColumnIndex': 0,
#                 'endColumnIndex': 3
#             },
#             'cell': {
#                 'userEnteredFormat': {
#                     'backgroundColor': color_theme
#                 }
#             },
#             'fields': 'userEnteredFormat.backgroundColor'
#         }
#     })

#     res = workbook.batch_update(body)

#     return res


if __name__ == '__main__':
    main()
