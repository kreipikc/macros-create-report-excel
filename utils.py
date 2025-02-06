import os
from datetime import timedelta, date
from typing import Optional, List
from num2words import num2words
from pydantic import ValidationError
from config import START_DATE
from schemas import ChecksDefault, TypeCheck


def validate_check(checks: List[ChecksDefault]) -> None:
    """Validate a list of checks to ensure all required fields are filled based on the check type.

    Args:
        checks (List[ChecksDefault]): A list of ChecksDefault objects to be validated.

    Returns:
        None

    Raises:
        Exception: If a required field is missing for a specific check type.
    """
    for check in checks:
        if check.type == TypeCheck.daily_allowance:
            if check.comment is None:
                raise Exception(f"There can't be a None comment for a '{check.type.value}' one.")
        else:
            if check.id_check is None or check.date is None:
                raise Exception(f"For the '{check.type.value}' type, the details must be filled in")


def create_check(check: tuple) -> Optional[ChecksDefault]:
    """Create a ChecksDefault object from a tuple of check data.

    Args:
        check (tuple): A tuple containing check data

    Returns:
        Optional[ChecksDefault]: A ChecksDefault object if the check data is valid, otherwise None.

    Raises:
        ValidationError: If the check data is invalid.
        Exception: If any other error occurs during the creation of the check.
    """
    try:
        check_res = ChecksDefault(
            number_str=int(check[0]),
            id_check=int(check[1]) if check[1] is not None else None,
            date=convert_excel_date_to_normal(int(check[2])) if check[2] is not None else None,
            sum_check=float(check[3]),
            type=TypeCheck(check[4].lower()),
            counterparty=check[5],
            counterparty_initials=check[6],
            counterparty_post=check[7],
            meeting_place=check[8],
            topic=check[9],
            comment=check[10],
        )
        return check_res
    except ValidationError as e:
        print(f"Validation error: {e}")
        return None
    except Exception as e:
        print(f"Error: {e}")
        return None


def convert_excel_date_to_normal(excel_date: int) -> date:
    """Convert an Excel date to a normal date.

    Args:
        excel_date (int): The date in Excel format.

    Returns:
        date: The converted date in normal format.
    """
    return START_DATE + timedelta(days=excel_date)


def sum_money_all_checks(checks: List[ChecksDefault]) -> float:
    """Calculate the total sum of all checks.

    Args:
        checks (List[ChecksDefault]): A list of ChecksDefault objects.

    Returns:
        float: The total sum of all checks.
    """
    sum_checks = 0.0
    for check in checks:
        sum_checks += check.sum_check
    return sum_checks


def convert_to_words(rubles: int, kopecks: int) -> str:
    """Convert a monetary amount to words in Russian.

    Args:
        rubles (int): The amount in rubles.
        kopecks (int): The amount in kopecks.

    Returns:
        str: The monetary amount in words.
    """
    rubles_in_words = num2words(rubles, lang='ru')
    # kopecks_in_words = num2words(kopecks, lang='ru')

    result = f"{rubles_in_words.capitalize()} рублей {kopecks} копеек ({rubles} руб. {kopecks} коп.)"
    return result

def get_absolute_path(relative_path: str) -> str:
    """Get the absolute path based on the location of the script.

    Args:
        relative_path (str): The relative path to the file.

    Returns:
        str: The absolute path to the file.
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(script_dir, relative_path)