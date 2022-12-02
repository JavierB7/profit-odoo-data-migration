import re
import pandas as pd
import os
import shutil

# The file "ids_odoo_for_update_costs_margins" should have id, default_code, alt_standard_price,
# pricelist_item_type_1_margin, pricelist_item_type_2_margin, pricelist_item_type_3_margin, pricelist_item_type_4_margin
# pricelist_item_type_1_freight, pricelist_item_type_2_freight, pricelist_item_type_3_freight, pricelist_item_type_4_freight

df_id_code = pd.read_csv(
    'files/ids_odoo_for_update_costs_margins.csv', delimiter=',', dtype=str, header='infer')
codes_of_odoo = df_id_code["default_code"].tolist()

df_cs_costs = pd.read_csv('files/costos_margenes_cs.csv', delimiter=',', dtype=str, header='infer')
codes_of_profit = df_id_code["default_code"].tolist()

products = {
    "id": [],
    "default_code": [],
    "alt_standard_price": [],
    "pricelist_item_type_1_margin": [],
    "pricelist_item_type_2_margin": [],
    "pricelist_item_type_3_margin": [],
    "pricelist_item_type_4_margin": [],
    "pricelist_item_type_1_freight": [],
    "pricelist_item_type_2_freight": [],
    "pricelist_item_type_3_freight": [],
    "pricelist_item_type_4_freight": [],
}

for code in codes_of_odoo:
    if code in codes_of_profit:
        match_code = df_cs_costs[df_cs_costs["default_code"] == code]
        if not match_code.empty:
            products["id"].append(df_id_code[df_id_code["default_code"] == code].iloc[0]["id"])
            products["default_code"].append(code)
            products["alt_standard_price"].append(match_code.iloc[0]["alt_standard_price"])
            products["pricelist_item_type_1_margin"].append(match_code.iloc[0]["profit_margin_1"])
            products["pricelist_item_type_2_margin"].append(match_code.iloc[0]["profit_margin_2"])
            products["pricelist_item_type_3_margin"].append(match_code.iloc[0]["profit_margin_3"])
            products["pricelist_item_type_4_margin"].append(match_code.iloc[0]["profit_margin_4"])
            products["pricelist_item_type_1_freight"].append(match_code.iloc[0]["freight_1"])
            products["pricelist_item_type_2_freight"].append(match_code.iloc[0]["freight_2"])
            products["pricelist_item_type_3_freight"].append(match_code.iloc[0]["freight_3"])
            products["pricelist_item_type_4_freight"].append(match_code.iloc[0]["freight_4"])

new_df = pd.DataFrame(products)
print(new_df)
new_df.to_csv("file_to_update.csv", encoding='utf-8', index=False)