# -*- coding: utf-8 -*-
"""
Created on Wed May 28 15:40:15 2025

@author: victor.bontemps
"""

from datetime import date
import pandas as pd
import matplotlib.pyplot as plt

df_clean = pd.DataFrame()
df = pd.read_json('data/data.json')

"""We clean the columns with specific format for the output"""

#"Barriers" is a list of dictionaries with two keys: "Frequency" and "PercentValue"
df["Frequency_"] = df["Barriers"].apply(
    lambda x: x[0].get("Frequency") if isinstance(x, list) and len(x) > 0 else None
df["PercentValue_"] = df["Barriers"].apply(
    lambda x: x[0].get("PercentValue") if isinstance(x, list) and len(x) > 0 else None
)
#"Identifiers" is a dictionary with two keys: "ISINs" and "InstrumentId"
df["ISIN_"] = df["Identifiers"].apply(
    lambda x: x{"ISINs"}[0] if isinstance(x, dict) and "ISINs" in x else None
)
#"Issuers" is a list of dictionaries with one key: "GroupName"
df["Issuer_"] = df["Issuers"].apply(
    lambda x: x[0].get("GroupName") if isinstance(x, list) and len(x) > 0 else None
)
#"Markets" is a list of dictionaries with one key: "Code"
df["Market_"] = df["Markets"].apply(
    lambda x: x[0].get("Code") if isinstance(x, list) and len(x) > 0 else None
)
#"Underlyings" is a list of dictionaries with one key: "Name"
df["Underlying_"] = df["Underlyings"].apply(
    lambda x: x[0].get("Name") if isinstance(x, list) and len(x) > 0 else None
)
#"Underlyings" is a list of dictionaries with one key: "Name"
df["Coupon"] = df["Underlyings"].apply(
    lambda x: x[0].get("Name") if isinstance(x, list) and len(x) > 0 else None
)

df_clean = df['Id', #OK
            'Type', #OK
            'InitialStrikeDateUTC', #OK
            'MaturityDateUTC', #OK
            'Tenor', #OK
            'Barriers', #--> [0]{'Frequency} & [0]{'PercentValue}
            'ProductCurrency', #OK
            'Identifiers', #-->{'ISINs}[0]
            'Issuers', #--> [0]{'GroupName}
            'Markets', #--> [0]{'Code'}
            'Categories', #OK
            'ProductGroup', #OK
            'PayoffStyles', #OK
            'DistributionChannels', #OK
            'Underlyings',#--> [0]{'Name'}
            'Coupons',
            'CapitalProtection',
            'SumMarketSalesVolume',
            'Name'] #OK


df.head(5).to_excel('data.xlsx', index=False)