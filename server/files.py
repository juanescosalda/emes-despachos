import os
import time
import logging
from pathlib import Path
from pdfminer.high_level import extract_text


class PDFManager:

    EMPTY = ''

    FIELDS = {
        'remision': "",
        'tienda': "",
        'direccion': "",
        'sector': "",
        'cliente': "",
        'telefono': "",
        'nit': "",
        'valor': "",
        "items": "",
        "notas": "",
        "saldo": ""
    }

    def __init__(self, folder_id: str) -> None:
        """
        Constructor

        Args:
            folder_id (str): Local folder path to Drive folder
        """
        self.old_pdfs = list()

        self.__path = \
            fr'G:\.shortcut-targets-by-id\{folder_id}\Despachos\Remisiones'

    def __get_tienda(self, ss_text: str) -> str:
        """
        Args:
            ss_text (str): extracted text

        Returns:
            str: store name
        """
        try:
            key_start = "Señores"
            key_end = "Atención"
            ind1 = ss_text.index(key_start)
            ind2 = ss_text.index(key_end)
            tienda_text = ss_text[ind1+len(key_start)+2:ind2-1]

            if tienda_text[-1] == "\n":
                tienda_text = tienda_text[:-1]

            tienda_text = tienda_text.replace('\n', ' - ')
            tienda_text = tienda_text.replace('--', '')
            tienda_text = tienda_text.strip()

            return tienda_text

        except:
            logging.error(
                'Error extrayendo nombre de la tienda',
                exc_info=True
            )

            return PDFManager.EMPTY

    def __get_telefono(self, ss_text: str) -> str:
        """
        Args:
            ss_text (str): extracted text

        Returns:
            str: phone number
        """
        try:
            key_start = "Tel."
            ind1 = ss_text.index(key_start)
            tel_text = ss_text[ind1+len(key_start)+1:]
            ind2 = tel_text.index("\n")
            tel_text = tel_text[:ind2]
            tel_text = tel_text.strip()

            return tel_text

        except:
            logging.error(
                'Error extrayendo teléfono de la tienda',
                exc_info=True
            )

            return PDFManager.EMPTY

    def __get_cliente(self, ss_text: str) -> str:
        """
        Args:
            ss_text (str): extracted text

        Returns:
            str: client name
        """
        try:
            key_start = "Atención :"
            ind1 = ss_text.index(key_start)
            cliente_text = ss_text[ind1+len(key_start)+1:]
            ind2 = cliente_text.index("\n")
            cliente_text = cliente_text[:ind2]
            cliente_text = cliente_text.strip()

            return cliente_text

        except:
            logging.error(
                'Error extrayendo nombre del cliente',
                exc_info=True
            )

            return PDFManager.EMPTY

    def __get_nit(self, ss_text: str) -> str:
        """
        Args:
            ss_text (str): extracted text

        Returns:
            str: nit number
        """
        try:
            key_start = "Nit:"
            key_end = "Asesor:"
            ind1 = ss_text.index(key_start)
            ind2 = ss_text.index(key_end)
            nit_text = ss_text[ind1+len(key_start)+1:ind2-1]
            nit_text = nit_text.replace('\n', '')
            nit_text = nit_text.strip()

            return nit_text

        except:
            logging.error(
                'Error extrayendo el nit de la empresa',
                exc_info=True
            )

            return PDFManager.EMPTY

    def __get_direccion_sector(self, ss_text: str) -> str:
        """
        Args:
            ss_text (str): extracted text

        Returns:
            str: address sector
        """
        try:
            key_start = "Atención :"
            key_end = "Tel."
            ind1 = ss_text.index(key_start)
            ind2 = ss_text.index(key_end)
            aux_text = ss_text[ind1+len(key_start)+1:ind2-1]
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

        except:
            logging.error(
                'Error extrayendo el sector donde está ubicada la empresa',
                exc_info=True
            )

            return PDFManager.EMPTY

    def __get_remision(self, ss_text: str) -> str:
        """
        Args:
            ss_text (str): extracted text

        Returns:
            str: address sector
        """
        try:
            key_start = "Vencimiento:"
            key_end = "Señores:"
            ind1 = ss_text.index(key_start)
            ind2 = ss_text.index(key_end)
            aux_text = ss_text[ind1+len(key_start)+1:ind2-1]
            aux_list = aux_text.split('\n')
            current_index = 0

            while (aux_list[current_index].strip().isnumeric()) == False:
                current_index = current_index + 1

            remision = aux_list[current_index].strip()

            return remision

        except:
            logging.error(
                'Error extrayendo el número de la remisión',
                exc_info=True
            )

            return PDFManager.EMPTY

    def __get_valor(self, ss_text: str) -> str:
        """
        Args:
            ss_text (str): extracted text

        Returns:
            str: invoice value
        """
        try:
            ss_text = ss_text.replace('\n', ' ')
            key_start = "Pagar"
            ind1 = ss_text.find(key_start)
            aux_text = ss_text[ind1+len(key_start)+1:]

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
                aux_text = ss_text[ind1+len(key_start)+1:ind2]
                valor = aux_text.split(' ')[-3]
            else:
                valor = values[1]

            return valor.strip()

        except:
            logging.error(
                'Error extrayendo el valor de la factura',
                exc_info=True
            )

            return PDFManager.EMPTY

    def __get_items(self, ss_text: str) -> str:
        """
        Args:
            ss_text (str): extracted text

        Returns:
            str
        """
        try:
            ss_text = ss_text.replace('\n', ' ')
            key_start = "Nº Items"
            ind1 = ss_text.index(key_start)
            aux_text = ss_text[ind1+len(key_start)+1:]
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

        except:
            logging.error(
                'Error extrayendo el número de items',
                exc_info=True
            )

        return PDFManager.EMPTY

    def __get_notas(self, ss_text: str) -> str:
        """
        Args:
            ss_text (str): extracted text

        Returns:
            str
        """
        try:
            ss_text = ss_text.replace('\n', ' ')
            key_start = "Nº Items"
            ind1 = ss_text.index(key_start)
            aux_text = ss_text[ind1+len(key_start)+2:]
            key_start = "Notas:"
            key_end = "Autorizo"
            ind1 = aux_text.index(key_start)
            ind2 = aux_text.index(key_end)
            notas = aux_text[ind1+len(key_start)+1:ind2-1]
            notas = notas.strip()

            return notas

        except:
            logging.error(
                'Error extrayendo las notas',
                exc_info=True
            )

            return PDFManager.EMPTY

    def __get_saldo(self, ss_text: str) -> str:
        """
        Args:
            ss_text (str): extracted text

        Returns:
            str
        """
        try:
            ss_text = ss_text.replace('\n', ' ')
            key_start = "Saldo del Cliente"
            ind1 = ss_text.index(key_start)
            aux_text = ss_text[ind1+len(key_start)+1:]
            stop_index = 0

            while aux_text[stop_index] != " ":
                stop_index = stop_index + 1

            saldo = aux_text[:stop_index]

            return saldo

        except:
            logging.error(
                'Error extrayendo el saldo pendiente del cliente',
                exc_info=True
            )

            return PDFManager.EMPTY

    def __get_pdf_fields(self, filename: str) -> dict:
        """
        Args:
            filename (str): file name

        Returns:
            dict: get required fields
        """
        all_text = extract_text(filename)
        key_start = "Señores:"
        key_end = "Asesor:"
        ind1 = all_text.index(key_start)
        ind2 = all_text.index(key_end)
        ss_text = all_text[ind1:ind2+len(key_end)]
        current_fields = PDFManager.FIELDS.copy()
        current_fields["tienda"] = self.__get_tienda(ss_text)
        current_fields["telefono"] = self.__get_telefono(ss_text)
        current_fields["nit"] = self.__get_nit(ss_text)
        current_fields["cliente"] = self.__get_cliente(ss_text)
        sector, direccion = self.__get_direccion_sector(ss_text)
        current_fields["sector"] = sector
        current_fields["direccion"] = direccion
        current_fields["remision"] = self.__get_remision(all_text)
        current_fields["valor"] = self.__get_valor(all_text)
        current_fields["items"] = self.__get_items(all_text)
        current_fields["notas"] = self.__get_notas(all_text)
        current_fields["saldo"] = PDFManager.EMPTY

        return current_fields

    def _get_new_pdfs(self) -> list:
        """
        Returns:
            list: new pdfs
        """
        pdf_files_object = Path(self.__path).glob("*.pdf")
        pdf_files = list()

        for pdf in pdf_files_object:
            pdf_files.append(pdf.name)

        return pdf_files

    def delete_file(self, file: str) -> None:
        """
        Delete file once has been processed
        """
        path = self.__path + "\\" + file

        if os.path.isfile(path):
            os.remove(path)
        else:
            print(f"Error: {file} file not found")

    def update(self) -> dict:
        """
        Main method
        """
        time.sleep(1)
        new_pdfs = self._get_new_pdfs()
        set_old_pdfs = set(self.old_pdfs)
        set_new_pdfs = set(new_pdfs)
        pdf2process = list(set_new_pdfs - set_old_pdfs)

        pdf_fields = dict()

        for name in pdf2process:
            pdf_path = self.__path + "\\" + name
            pdf_values = self.__get_pdf_fields(pdf_path)
            pdf_fields.update({name: pdf_values})

        self.old_pdfs = new_pdfs

        return pdf_fields
