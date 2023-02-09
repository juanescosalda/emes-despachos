import firebase_admin
from firebase_admin import db
from datetime import datetime


class Database:

    def __init__(
            self,
            json_path: str,
            db_url: str) -> None:
        """
        Constructor

        Args:
            json_path (str): firebase json path
            db_url (str): database url

        Raises:
            Exception: unauthorized firebase admin session
        """
        self.cred_firebase = firebase_admin.credentials.Certificate(
            json_path
        )

        if not firebase_admin._apps:
            fs = firebase_admin.initialize_app(
                self.cred_firebase,
                {
                    'databaseURL': db_url
                }
            )

            if fs is None:
                raise Exception("Error creating Firebase object")
            else:
                self.ref = db.reference("/empaques/")

    def insert_values(
            self,
            items: int,
            factura: int,
            farmacia: str,
            valor: int) -> None:
        """
        Args:
            items (int): total products
            factura (int): invoice number
            farmacia (str): client name
            valor (int): total invoice value
        """
        self.ref.push().set(
            {
                "articulos": items,
                "bolsas": 0,
                "cajas": 0,
                "empacador": "",
                "estado": 0,
                "factura": factura,
                "farmacia": farmacia,
                "fecha": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "lios": 0,
                "region": "",
                "valor": valor,
                "uid": ""
            }
        )
