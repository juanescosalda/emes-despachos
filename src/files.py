import os
import re
import time
import logging
from typing import Dict, List, Any
from pathlib import Path
from constants import *
from pdfminer.high_level import extract_text


class PDFManager:

    FIELDS = \
        {
            'remision': "",
            'tienda': "",
            'direccion': "",
            'sector': "",
            'cliente': "",
            'telefono': "",
            'nit': "",
            'valor': "",
            "items": "",
            "facturador": "",
            "forma": "",
            "notas": "",
        }

    def __init__(self, folder_id: str) -> None:
        """
        Constructor

        Args:
            folder_id (str): Local folder path to Drive folder
        """
        self.__old_pdfs = []

        self.__path = \
            fr'G:\.shortcut-targets-by-id\{folder_id}\Despachos\Remisiones'

    def __get_tienda(self, ss_text: str) -> str:
        """
        Args:
            ss_text (str): Common text

        Returns:
            str: Name of "Tienda"
        """
        try:
            key_start = "Señores"
            key_end = "Atención"
            ind1 = ss_text.index(key_start)
            ind2 = ss_text.index(key_end)
            tienda_text = ss_text[ind1 + len(key_start) + 2:ind2 - 1]

            if tienda_text[-1] == "\n":
                tienda_text = tienda_text[:-1]

            tienda_text = tienda_text.replace('\n', ' - ')
            tienda_text = tienda_text.replace('--', '')
            tienda_text = tienda_text.strip()

            return tienda_text

        except Exception as e:
            logging.error(
                f'Error extrayendo el nombre de la tienda >>> {e}',
                exc_info=True
            )
            return EMPTY_STRING

    def __get_telefono(self, ss_text: str) -> str:
        """
        Args:
            ss_text (str): Common text

        Returns:
            str: Phone number
        """
        try:
            key_start = "Tel."
            ind1 = ss_text.index(key_start)
            tel_text = ss_text[ind1 + len(key_start) + 1:]
            ind2 = tel_text.index("\n")
            tel_text = tel_text[:ind2]
            tel_text = tel_text.strip()

            return tel_text

        except Exception as e:
            logging.error(
                f'Error {e} extrayendo teléfono de la tienda',
                exc_info=True
            )
            return EMPTY_STRING

    def __get_cliente(self, ss_text: str) -> str:
        """
        Args:
            ss_text (str): Common text

        Returns:
            str: Client name
        """
        try:
            key_start = "Atención :"
            ind1 = ss_text.index(key_start)
            cliente_text = ss_text[ind1 + len(key_start) + 1:]
            ind2 = cliente_text.index("\n")
            cliente_text = cliente_text[:ind2]
            cliente_text = cliente_text.strip()

            return cliente_text

        except Exception as e:
            logging.error(
                f'Error {e} extrayendo nombre del cliente',
                exc_info=True
            )
            return EMPTY_STRING

    def __get_nit(self, ss_text: str) -> str:
        """
        Args:
            ss_text (str): Common text

        Returns:
            str: NIT number
        """
        try:
            key_start = "Nit:"
            key_end = "Asesor:"
            ind1 = ss_text.index(key_start)
            ind2 = ss_text.index(key_end)
            nit_text = ss_text[ind1 + len(key_start) + 1:ind2 - 1]
            nit_text = nit_text.replace('\n', '')
            nit_text = nit_text.strip()

            return nit_text

        except Exception as e:
            logging.error(
                f'Error {e} extrayendo el nit de la empresa',
                exc_info=True
            )
            return EMPTY_STRING

    def __get_direccion_sector(self, ss_text: str) -> str:
        """
        Args:
            ss_text (str): Common text

        Returns:
            str: Client address
        """
        try:
            key_start = "Atención :"
            key_end = "Tel."
            ind1 = ss_text.index(key_start)
            ind2 = ss_text.index(key_end)
            aux_text = ss_text[ind1 + len(key_start) + 1:ind2 - 1]
            aux_list = aux_text.split('\n')
            current_index = len(aux_list) - 1

            while aux_list[current_index] == "":
                current_index = current_index - 1

            sector = aux_list[current_index]
            current_index = current_index - 1

            while aux_list[current_index] == "":
                current_index = current_index - 1

            direccion = aux_list[current_index]
            direccion = direccion.strip()
            sector = sector.strip()

            return sector, direccion

        except Exception as e:
            logging.error(
                f'Error {e} extrayendo el sector donde está ubicada la empresa',
                exc_info=True
            )
            return EMPTY_STRING

    def __get_remision(self, ss_text: str) -> str:
        """
        Args:
            ss_text (str): Common text

        Returns:
            str: Number code of "Remision"
        """
        try:
            key_start = "Vencimiento:"
            key_end = "Señores:"
            ind1 = ss_text.index(key_start)
            ind2 = ss_text.index(key_end)
            aux_text = ss_text[ind1 + len(key_start) + 1:ind2 - 1]
            aux_list = aux_text.split('\n')
            current_index = 0

            while (aux_list[current_index].strip().isnumeric()) == False:
                current_index = current_index + 1

            remision = aux_list[current_index].strip()

            return remision

        except Exception as e:
            logging.error(
                f'Error {e} extrayendo el número de la remisión',
                exc_info=True
            )
            return EMPTY_STRING

    def __get_valor(self, ss_text: str) -> str:
        """
        Args:
            ss_text (str): Common text

        Returns:
            str: Invoice value
        """
        try:
            ss_text = ss_text.replace('\n', ' ')
            key_start = "Pagar"
            ind1 = ss_text.find(key_start)
            aux_text = ss_text[ind1 + len(key_start) + 1:]

            values = aux_text.split(' ')

            if values[1] == '$':
                if values[4] == '$':
                    valor = values[7]
                else:
                    valor = values[4]
            elif values[0] == '' and values[1] == '':
                key_start = "Autor:   Cambio"
                ind2 = ss_text.find("Saldo del Cliente")
                ind1 = ss_text.find(key_start)
                aux_text = ss_text[ind1 + len(key_start) + 1:ind2]
                valor = aux_text.split(' ')[-3]
            else:
                valor = values[1]

            return valor.strip()

        except Exception as e:
            logging.error(
                f'Error {e} extrayendo el valor de la factura',
                exc_info=True
            )
            return EMPTY_STRING

    def __get_facturador(self, ss_text: str) -> str:
        """
        Args:
            ss_text (str): Common text

        Returns:
            str: Invoice biller name
        """
        try:
            ss_text = ss_text.replace('\n', ' ')
            pattern = r'Facturador\s+([A-Z\s]+)'

            # Search for the name of the biller in the text
            res = re.search(pattern, ss_text)

            if res:
                biller_name = res.group(1)
                biller_name = re.sub(r'\s+[A-Z]$', '', biller_name)

                return biller_name

            return EMPTY_STRING

        except Exception as e:
            logging.error(
                f'Error {e} extrayendo el valor de la factura',
                exc_info=True
            )
            return EMPTY_STRING

    def __get_forma_pago(self, ss_text: str) -> str:
        """
        Args:
            ss_text (str): Common text

        Returns:
            str: Payment method
        """
        try:
            ss_text = ss_text.replace('\n', ' ')

            key_start = 'Forma de Pago: '

            ind1 = ss_text.find(key_start)
            ind2 = ss_text.find('Elaboró')

            if ind2 - ind1 > 50:
                ind2 = ss_text.find('RECIBO DE CAJA PROVISIONAL')

            if ind1 == -1 or ind2 == -1:
                return EMPTY_STRING
            else:
                forma_pago = ss_text[ind1 + len(key_start):ind2 - 2]
                forma_pago = forma_pago.lstrip()
                return forma_pago

        except Exception as e:
            logging.error(
                f'Error {e} extrayendo la forma de pago',
                exc_info=True
            )
            return EMPTY_STRING

    def __get_items(self, ss_text: str) -> str:
        """
        Args:
            ss_text (str): Common text

        Returns:
            str (str): Number of items
        """
        try:
            ss_text = ss_text.replace('\n', ' ')
            key_start = "Nº Items"
            ind1 = ss_text.index(key_start)
            aux_text = ss_text[ind1 + len(key_start) + 1:]
            start_index = 0

            while aux_text[start_index] == " ":
                start_index = start_index + 1

            stop_index = start_index

            while aux_text[stop_index] != " ":
                stop_index = stop_index + 1

            items = aux_text[start_index:stop_index]
            aux_items = items.split(".")
            items = aux_items[0]

            return items

        except Exception as e:
            logging.error(
                f'Error {e} extrayendo el número de items',
                exc_info=True
            )
            return EMPTY_STRING

    def __get_notas(self, ss_text: str) -> str:
        """
        Args:
            ss_text (str): Common text

        Returns:
            str (str): Notas
        """
        try:
            ss_text = ss_text.replace('\n', ' ')
            key_start = "Nº Items"
            ind1 = ss_text.index(key_start)
            aux_text = ss_text[ind1 + len(key_start) + 2:]
            key_start = "Notas:"
            key_end = "Autorizo"
            ind1 = aux_text.index(key_start)
            ind2 = aux_text.index(key_end)
            notas = aux_text[ind1 + len(key_start) + 1:ind2 - 1]
            notas = notas.strip()

            return notas

        except Exception as e:
            logging.error(
                f'Error {e} extrayendo las notas',
                exc_info=True
            )
            return EMPTY_STRING

    def __get_pdf_fields(
            self,
            filename: str) -> Dict[str, Any] | None:
        """
        Args:
            filename (str): File name

        Returns:
            Dict[str, Any]: Required fields
        """
        try:
            all_text = extract_text(filename)
            key_start = "Señores:"
            key_end = "Asesor:"
            ind1 = all_text.index(key_start)
            ind2 = all_text.index(key_end)
            ss_text = all_text[ind1:ind2 + len(key_end)]

            # Copy dict structure
            fields = PDFManager.FIELDS.copy()

            # Fill dict fields
            fields["tienda"] = self.__get_tienda(ss_text)
            fields["telefono"] = self.__get_telefono(ss_text)
            fields["nit"] = self.__get_nit(ss_text)
            fields["cliente"] = self.__get_cliente(ss_text)
            sector, direccion = self.__get_direccion_sector(ss_text)
            fields["sector"] = sector
            fields["direccion"] = direccion
            fields["remision"] = self.__get_remision(all_text)
            fields["valor"] = self.__get_valor(all_text)
            fields["facturador"] = self.__get_facturador(all_text)
            fields["forma"] = self.__get_forma_pago(all_text)
            fields["items"] = self.__get_items(all_text)
            fields["notas"] = self.__get_notas(all_text)

            return fields

        except Exception as e:
            logging.error(
                f"Error al extraer los datos necesarios del PDF >>> {e}",
                exc_info=True
            )
            return None

    def __get_new_pdfs(self) -> List[str] | None:
        """
        Returns:
            List[str]: New PDFs added
        """
        try:
            pdf_files_object = Path(self.__path).glob("*.pdf")
            pdf_files = \
                [
                    pdf.name
                    for pdf in pdf_files_object
                ]

            return pdf_files

        except Exception as e:
            logging.error(
                f"Error al obtener la lista de nuevos PDFs >>> {e}",
                exc_info=True
            )
            return None

    def delete_file(self, file: str) -> None:
        """
        Delete file once has been processed

        Args:
            file (str): File name
        """
        try:
            path = self.__path + "\\" + file

            if os.path.isfile(path):
                os.remove(path)
            else:
                logging.info(f"Error: {file} file not found")

        except Exception as e:
            logging.error(
                f"Error al intentar eliminar el archivo: {file} >>> {e}",
                exc_info=True
            )

    def delete_all_files(self) -> None:
        """
        Delete all non-pdf files from source folder
        """
        try:
            for file in os.listdir(self.__path):
                if file.endswith(".rpt") or file.endswith(".pdf"):
                    # Constructs the complete file path
                    file_path = os.path.join(self.__path, file)
                    # Delete file
                    os.remove(file_path)

        except Exception as e:
            logging.error(
                f"Error al intentar eliminar todos los archivos del folder de Remisiones >>> {e}",
                exc_info=True
            )

    def update(self) -> Dict[str, Any] | None:
        """
        Main method
        """
        try:
            time.sleep(1)

            new_pdfs = self.__get_new_pdfs()
            set_old_pdfs = set(self.__old_pdfs)
            set_new_pdfs = set(new_pdfs)
            pdf2process = list(set_new_pdfs - set_old_pdfs)

            pdf_fields = {}

            for name in pdf2process:
                pdf_path = self.__path + "\\" + name
                pdf_values = self.__get_pdf_fields(pdf_path)

                if pdf_values is not None:
                    pdf_fields.update({name: pdf_values})

            self.__old_pdfs = new_pdfs

            return pdf_fields

        except Exception as e:
            logging.error(
                f"Error al obtener los datos de los nuevos PDFs >>> {e}",
                exc_info=True
            )
            return None
