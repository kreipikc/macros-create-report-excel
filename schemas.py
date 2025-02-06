from datetime import date
from enum import Enum
from typing import Optional
from pydantic import BaseModel


class TypeCheck(Enum):
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


class ChecksDefault(BaseModel):
    number_str: int
    id_check: Optional[int]
    date: Optional[date]
    sum_check: float
    type: TypeCheck
    counterparty: Optional[str]
    counterparty_initials: Optional[str]
    counterparty_post: Optional[str]
    meeting_place: Optional[str]
    topic: Optional[str]
    comment: Optional[str]


class AdditionalInfo(BaseModel):
    employee: str
    report_month: date
    post: str