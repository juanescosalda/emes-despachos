

def rgb2p(code: int) -> float:
    return code / 255.0


def str2int(val: str) -> int:
    try:
        val = val.strip("$")
        indPunto = val.find(".")
        val = val[:indPunto]
        val = val.replace(",", "")
        val = int(val)
    except:
        return int(0)

    return val


def str_to_int(val: str) -> int:
    try:
        val = val.strip("$")
        val = val[:len(val)]
        val = val.replace(",", "")
        val = int(val.strip())
    except:
        return int(0)

    return val


def is_dict(obj) -> bool:
    """
    Checks if object is dict

    Returns:
        bool: true if all objects are dicts
    """
    return isinstance(obj, dict)


def all_dict(*args) -> bool:
    """
    Checks if all objects are dicts

    Returns:
        bool: true if all objects are dicts
    """
    return all(is_dict(obj) for obj in args)


def safe_int(value: str) -> int:
    """
    Safe mode to convert 

    Args:
        value (str): Value as string

    Returns:
        int: Converted value to integer
    """
    try:
        return int(value) if value != '' else 0
    except ValueError:
        return 0
