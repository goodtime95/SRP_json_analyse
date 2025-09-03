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

#             ['Id', #OK
#             'Type', #OK
#             'InitialStrikeDateUTC', #OK
#             'MaturityDateUTC', #OK
#             'Tenor', #OK
#             'Barriers', #--> [0]{'Frequency} & [0]{'PercentValue}
#             'ProductCurrency', #OK
#             'Identifiers', #-->{'ISINs}[0]
#             'Issuers', #--> [0]{'GroupName}
#             'Markets', #--> [0]{'Code'}, [0]{'Distributors'}, [0]{'Brochures'}    
#             'Categories', #OK
#             'ProductGroup', #OK
#             'PayoffStyles', #OK
#             'DistributionChannels', #OK
#             'AssetClasses', --> [0]{'Name'}
#             'Underlyings',#--> #[0]{'Name'}
#             'Autocalls',#--> #[0]{'DateUTC'} + [0]{'Level'} + [0]{'Payout'} for 0 and -1
#             'Coupons', #--> #[0]{'MaxCoupon'} + [0]{'MinCoupon'}
#             'CapitalProtection', #OK
#             'SumMarketSalesVolume',#-->'Type': 'Estimated', 'IsPublic': True, 'Amounts'['Native']['Value']
#             'Name'] #OK

"""We clean the columns with specific format for the output"""

#"Barriers" --> PDI_Type_ & PDI_Barrier_
df["PDI_Type_"] = df["Barriers"].apply(
    lambda x: x[0].get("Frequency") if isinstance(x, list) and len(x) > 0 else None
)
df["PDI_Barrier_"] = df["Barriers"].apply(
    lambda x: x[0].get("PercentValue") if isinstance(x, list) and len(x) > 0 else None
)

#"Identifiers" --> ISIN_
df["ISIN_"] = df["Identifiers"].apply(
    lambda x: x.get("ISINs",[None])[0] if x is not None and x.get("ISINs") else None
)

#"Issuers" --> Issuer_
df["Issuer_"] = df["Issuers"].apply(
    lambda x: x[0].get("GroupName") if isinstance(x, list) and len(x) > 0 else None
)

#"Markets" --> Country_, Distributor_, FT_
df["Country_"] = df["Markets"].apply(
    lambda x: x[0].get("Code") if isinstance(x, list) and len(x) > 0 else None
)
df["Distributor_"] = df["Markets"].apply(
    lambda x: x[0].get("Distributors")[0].get("Name") if isinstance(x, list) and len(x) > 0 else None
)
df["FT_"] = df["Markets"].apply(
    lambda x: x[0].get("Brochures",[None])[0].get("DownloadUri",None) if isinstance(x, list) and len(x) > 0 and x[0].get("Brochures") and len(x[0].get("Brochures")) > 0 else None
)

#"AssetClasses" --> AssetClass_
df["AssetClass_"] = df["AssetClasses"].apply(
    lambda x: x[0].get("Name") if isinstance(x, list) and len(x) > 0 else None
)

#"Underlyings" --> Underlying_
df["Underlying_"] = df["Underlyings"].apply(
    lambda x: x[0].get("Name") if isinstance(x, list) and len(x) > 0 else None
)


# #"Underlyings" is a list of dictionaries with one key: "Name"
# df["Autocall_First_Date"] = df["Underlyings"].apply(
#     lambda x: x[0].get("Name") if isinstance(x, list) and len(x) > 0 else None
# )
# #"Underlyings" is a list of dictionaries with one key: "Name"
# df["Autocall_First_Payout"] = df["Underlyings"].apply(
#     lambda x: x[0].get("Name") if isinstance(x, list) and len(x) > 0 else None
# )
# #"Underlyings" is a list of dictionaries with one key: "Name"
# df["Autocall_Last_Date"] = df["Underlyings"].apply(
#     lambda x: x[0].get("Name") if isinstance(x, list) and len(x) > 0 else None
# )
# #"Underlyings" is a list of dictionaries with one key: "Name"
# df["Autocall_Last_Payout"] = df["Underlyings"].apply(
#     lambda x: x[0].get("Name") if isinstance(x, list) and len(x) > 0 else None
# )
# #"Underlyings" is a list of dictionaries with one key: "Name"
# df["Coupon"] = df["Underlyings"].apply(
#     lambda x: x[0].get("Name") if isinstance(x, list) and len(x) > 0 else None
# )

# df_clean = df['Id', #OK
#             'Type', #OK
#             'InitialStrikeDateUTC', #OK
#             'MaturityDateUTC', #OK
#             'Tenor', #OK
#             'Barriers', #--> [0]{'Frequency} & [0]{'PercentValue}
#             'ProductCurrency', #OK
#             'Identifiers', #-->{'ISINs}[0]
#             'Issuers', #--> [0]{'GroupName}
#             'Markets', #--> [0]{'Code'}
#             'Categories', #OK
#             'ProductGroup', #OK
#             'PayoffStyles', #OK
#             'DistributionChannels', #OK
#             'Underlyings',#--> #[0]{'Name'}
#             'Autocalls',#--> #[0]{'DateUTC'} + [0]{'Level'} + [0]{'Payout'} for 0 and -1
#             'Coupons', #--> #[0]{'MaxCoupon'} + [0]{'MinCoupon'}
#             'CapitalProtection', #OK
#             'SumMarketSalesVolume',#-->'Type': 'Estimated', 'IsPublic': True, 'Amounts'['Native']['Value']
#             'Name'] #OK

# test_dic =df['Autocalls'].iloc[0][0]
# print(test_dic.keys())

df.head(5).to_excel('data.xlsx', index=False)