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
df.to_excel('output/data0.xlsx', index=False)


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

#"Underlyings" --> Underlying_ & SectorName
df["Underlying_"] = df["Underlyings"].apply(
    lambda x: x[0].get("Name") if isinstance(x, list) and len(x) > 0 else None
)
df["Underlying_Type_"] = df["Underlyings"].apply(
    lambda x: x[0].get("SectorName") if isinstance(x, list) and len(x) > 0 else None
)

#"Autocalls" --> AC_LastDate_, AC_LastLevel_, AC_LastPayout_, AC_FirstDate_, AC_FirstLevel_, AC_FirstPayout_ 
df["AC_LastDate_"] = df["Autocalls"].apply(
    lambda x: x[0].get("DateUTC") if isinstance(x, list) and len(x) > 0 else None
)
df["AC_LastLevel_"] = df["Autocalls"].apply(
    lambda x: x[0].get("Level") if isinstance(x, list) and len(x) > 0 else None
)
df["AC_LastPayout_"] = df["Autocalls"].apply(
    lambda x: x[0].get("Payout") if isinstance(x, list) and len(x) > 0 else None
)
df["AC_FirstDate_"] = df["Autocalls"].apply(
    lambda x: x[-1].get("DateUTC") if isinstance(x, list) and len(x) > 0 else None
)
df["AC_FirstLevel_"] = df["Autocalls"].apply(
    lambda x: x[-1].get("Level") if isinstance(x, list) and len(x) > 0 else None
)
df["AC_FirstPayout_"] = df["Autocalls"].apply(
    lambda x: x[-1].get("Payout") if isinstance(x, list) and len(x) > 0 else None
)

#"Coupons" --> MinCoupon_, MaxCoupon_
df["MinCoupon_"] = df["Coupons"].apply(
    lambda x: x[0].get("MinCoupon") if isinstance(x, list) and len(x) > 0 else None
)
df["MaxCoupon_"] = df["Coupons"].apply(
    lambda x: x[0].get("MaxCoupon") if isinstance(x, list) and len(x) > 0 else None
)
#"Wrappers" --> Wrapper_
df["Wrapper_"] = df["Wrappers"].apply(
    lambda x: x[0].get("Name") if isinstance(x, list) and len(x) > 0 else None
)

#"AutoCallFrequency" --> AutoCallFreq_
df["AutoCallFreq_"] = df["AutoCallFrequency"].apply(
    lambda x: x[0] if isinstance(x, list) and len(x) > 0 else None
)

#"SumMarketSalesVolume" --> Volume_
df["Volume_"] = df["SumMarketSalesVolume"].apply(
    lambda x: x.get("Amounts").get("Native").get("Value") if isinstance(x, dict) and x.get("Amounts") and x.get("Amounts").get("Native") else None
)

#"Descriptions" --> Description_
df["Description_"] = df["Descriptions"].apply(
    lambda x: x[0].get("Value") if isinstance(x, list) and len(x) > 0 else None
)

#"PotentialMaxPayout" --> MaxPayout_
df["MaxPayout_"] = df["PotentialMaxPayout"].apply(
    lambda x: x.get("MaxAnnualized") if isinstance(x, dict) and x.get("MaxAnnualized") else None
)
#All fields in df_clean
fields_to_keep_clean = ["Id",
                        "Type",
                        "InitialStrikeDateUTC",
                        "MaturityDateUTC",
                        "Tenor",
                        "PDI_Type_",
                        "PDI_Barrier_",
                        "ProductCurrency",
                        "ISIN_",
                        "Issuer_",
                        "Country_",
                        "Distributor_",
                        "FT_",
                        "Categories",
                        "ProductGroup",
                        "PayoffStyles",
                        "AssetClass_",
                        "AssetClass_",
                        "Underlying_",
                        "Underlying_Type_",
                        "AC_LastDate_",
                        "AC_LastLevel_",
                        "AC_LastPayout_",
                        "AC_FirstDate_",
                        "AC_FirstLevel_",
                        "AC_FirstPayout_",
                        "MinCoupon_",
                        "MaxCoupon_",
                        "MaxPayout_",
                        "Wrapper_",
                        "AutoCallFreq_",
                        "Volume_",
                        "CapitalProtection",
                        "Name",
                        "Description_"]

df_clean = df[fields_to_keep_clean]
df_clean.to_excel('output/data_clean.xlsx', index=False)

#Interest Rates Products
df_ir_products = df[df["AssetClass_"] == "Interest Rate"]
fields_to_keep_ir_products = ["Country_",
                            "Name",
                            "Issuer_",
                            "Distributor_",
                            "Wrapper_",
                            "Volume_",
                            "ProductCurrency",
                            "Underlying_",
                            "InitialStrikeDateUTC",
                            "Tenor",
                            "CapitalProtection",
                            "MinCoupon_",
                            "MaxCoupon_",
                            "MaxPayout_",
                            "AutoCallFreq_",
                            "AC_FirstDate_",
                            "AC_FirstLevel_",
                            "AC_FirstPayout_",
                            "AC_LastLevel_",
                            "AC_LastPayout_",
                            "Description_",
                            "ISIN_",
                            "FT_"
                            ]

df_ir_products = df_ir_products[fields_to_keep_ir_products]
df_ir_products.to_excel('output/data_ir_products.xlsx', index=False)

#Credit Products
df_credit_products = df[df["AssetClass_"] == "Credit"]
fields_to_keep_credit_products = ["Country_",
                            "Name",
                            "Issuer_",
                            "Distributor_",
                            "Wrapper_",
                            "Volume_",
                            "ProductCurrency",
                            "Underlying_",
                            "InitialStrikeDateUTC",
                            "Tenor",
                            "MaxPayout_",
                            "Description_",
                            "ISIN_",
                            "FT_"
                            ]

df_credit_products = df_credit_products[fields_to_keep_credit_products]
df_credit_products.to_excel('output/data_credit_products.xlsx', index=False)



