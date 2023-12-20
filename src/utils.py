'''
Common useful functions
'''
import locale
import logging
from typing import Dict, Any

locale.setlocale(locale.LC_ALL, '')


def rgb2p(code: int) -> float:
    """
    Convert an RGB color code to a normalized float value.

    Args:
        code (int): RGB color code (0-255).

    Returns:
        float: Normalized value between 0 and 1.
    """
    return code / 255.0


def str2int(text: str) -> int:
    """
    Safely convert a string to an integer.

    Args:
        text (str): Input string.

    Returns:
        int: Converted integer value. If the conversion fails, returns 0.
    """
    try:
        out_value = 0

        if text:
            value = text.strip("$")
            point_index = value.find(".")
            value = value[:point_index]
            out_value = int(value.replace(",", "").strip())

        return out_value

    except Exception as e:
        logging.exception(
            f"Converting str to integer => Error: {e}",
            exc_info=True
        )
        return 0


def str_to_int(text: str) -> int:
    """
    Safely convert a string to an integer.

    Args:
        text (str): Input string.

    Returns:
        int: Converted integer value. If the conversion fails, returns 0.
    """
    try:
        out_value = 0

        if text:
            value = text.strip("$")
            value = value[:len(value)]
            value = value.replace(",", "").strip()
            out_value = int(value) if value else 0

        return out_value

    except Exception as e:
        logging.exception(
            f"Converting str to integer => Error: {e} => Text: {text}",
            exc_info=True
        )
        return 0


def is_dict(obj) -> bool:
    """
    Checks if object is dict

    Returns:
        bool: True if all objects are dicts
    """
    return isinstance(obj, dict)


def all_dict(*objs) -> bool:
    """
    Checks if all objects are dicts

    Returns:
        bool: True if all objects are dicts
    """
    return all(is_dict(obj) for obj in objs)


def safe_int(value: str) -> int:
    """
    Safe mode to convert 

    Args:
        value (str): Value as string

    Returns:
        int: Converted value to integer
    """
    try:
        return int(value) if value else 0
    except ValueError:
        return -1


def str_to_bool(text: str) -> bool:
    """
    Convert a string into a boolean value.

    Args:
        text (str): Boolean as a string.

    Returns:
        bool: Boolean value. True if the string is 'TRUE', False otherwise.
    """
    return text.upper() == 'TRUE' if isinstance(text, str) else False


def has_value(
        data_dict: Dict[Any, Any],
        key: str,
        value: Any):
    """
    Check if a dictionary has a specific key with a specified value.

    Args:
        data_dict (Dict[Any, Any]): Dictionary to check.
        key (str): Key to check.
        value (Any): Value to compare.

    Returns:
        bool: True if the dictionary has the key and the value matches, False otherwise.
    """
    return key in data_dict and data_dict[key] == value


def to_currency(value: Any) -> str:
    """
    Convert a value to a currency-formatted string.

    Args:
        value (Any): Value to convert.

    Returns:
        str: Currency-formatted string. 
             If the value is a string, it is converted to a float and then formatted. 
             If the conversion fails, returns "$0.0".
    """
    try:
        val = 0.0 if isinstance(value, str) else float(value)
        return locale.currency(val, grouping=True)

    except Exception as e:
        logging.error(
            f"Error al convertir el valor de float/int a string >>> {e}"
        )
        return "$0.0"


def from_currency(text: str) -> int:
    """
    Convert a currency-formatted string to an integer value.

    Args:
        text (str | None): Currency-formatted string.

    Returns:
        int: Converted integer value. If the conversion fails or the string is empty, returns -1.
    """
    try:
        if not text:
            return 0

        # Delete "peso" symbol and non-numerical characters
        text = text.replace('$', '').replace(',', '').strip()
        return int(float(text))

    except Exception as e:
        logging.error(
            f"Error al convertir el valor de string a entero >>> {e}"
        )
        return -1
