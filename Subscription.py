from dataclasses import dataclass
from abc import ABC
from Settings import settings
import datetime


@dataclass
class Subscription(ABC):
    price: float
    free_clicks: int
    extra_clicks: int
    date: datetime.date

    def __post_init__(self):
        self.total_clicks = self.free_clicks + self.extra_clicks
        if self.price == settings.light_fee:
            self.subscription_type = 'Light'
            self.allowed_clicks = settings.light_clicks
            self.extra_fee = settings.light_extra_fee
        elif self.price == settings.standard_fee:
            self.subscription_type = 'Standard'
            self.allowed_clicks = settings.standard_clicks
            self.extra_fee = settings.standard_extra_free
        elif self.price == settings.premium_fee:
            self.subscription_type = 'Premium'
            self.allowed_clicks = settings.premium_clicks
            self.extra_fee = settings.premium_extra_fee

    def change_single_subscription(self, sub_type: str):
        self.subscription_type = sub_type
        if sub_type == 'Light':
            self.price = settings.light_fee
            self.allowed_clicks = settings.light_clicks
            self.extra_fee = settings.light_extra_fee
        elif sub_type == 'Standard':
            self.price = settings.standard_fee
            self.allowed_clicks = settings.standard_clicks
            self.extra_fee = settings.standard_extra_free
        elif sub_type == 'Premium':
            self.price = settings.premium_fee
            self.allowed_clicks = settings.premium_clicks
            self.extra_fee = settings.premium_extra_fee
        else:
            self.extra_fee = 0
            self.allowed_clicks = 0
            self.subscription_type = None
            self.price = 0

    def change_single_subscription_with_parameters(self, data: dict):
        self.subscription_type = data['Type']
        self.price = data['Price']
        self.allowed_clicks = data['Allowed clicks']
        self.extra_fee = data['Extra fee']
