from datetime import date
from enum import Enum
from typing import Optional, Union
from pydantic import BaseModel


class TypeCheck(str, Enum):
    """ All type of checks """
    representative_offices_event = "представительские - мероприятие"
    representative_offices_present = "представительские - подарки"
    round_table_discussion_Club = "круглый стол/дискуссионный клуб"
    chancellery = "канцелярия"
    office_meetings = "служебные совещания"
    car_expenses = "расходы на автомобиль"
    fuel_and_lubricants = "гсм"
    other_business_trips = "прочие командировочные"
    daily_allowance = "суточные"


class TypeDocument(str, Enum):
    """ All type of documents """
    Cash_receipt_Representation_expenses = "Кассовый чек Представительские расходы"
    CCT_receipt = "Чек ККТ"
    Fuel_and_lubricants_cashiers_check = "Кассовый чек ГСМ"
    Parking_Report = "Отчет о парковке"
    Invoice = "Накладная"
    Cash_receipt_Daily_allowance = "Кассовый чек Суточные"


class ChecksDefault(BaseModel):
    number_str: int
    type_document: TypeDocument
    id_check: Optional[Union[int, str]]
    date: Optional[date]
    sum_check: float
    type: TypeCheck
    counterparty: Optional[str]
    counterparty_participant: Optional[Union[int, str]]
    counterparty_post: Optional[str]
    meeting_place: Optional[str]
    medication: Optional[str]
    topic: Optional[str]
    name_present: Optional[str]
    comment: Optional[str]


class AdditionalInfo(BaseModel):
    employee: str
    report_month: date
    date_report: date
    post: str
    department: str