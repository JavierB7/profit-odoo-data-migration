import re
import pandas as pd
import os
import shutil

def copy_img_to_upload():
    for code in codes_of_odoo_without_img:
        images = os.listdir('images')
        for img in images:
            if not img.startswith('.'):
                if code in img:
                    shutil.copyfile(f"images/{img}", f"images_to_upload/{img}")

df_id_code = pd.read_csv('codigos_migracion_imagenes.csv', delimiter=',', dtype=str, header='infer')
codes_of_odoo_without_img = df_id_code["default_code"].tolist()

df_images = pd.read_csv('imagenes_productos_masi_main.csv', delimiter=',', dtype=str, header='infer')
codes_of_profit_with_img = df_images["code"].tolist()

product_ids = []
product_img_urls = []
for code in codes_of_odoo_without_img:
    if code in codes_of_profit_with_img:
        product_id = df_id_code[df_id_code["default_code"] == code].iloc[0]["id"]
        img_url = df_images[df_images["code"] == code].iloc[0]["url"]
        product_ids.append(product_id)
        product_img_urls.append(img_url)

data = {
    "id": product_ids,
    "image_1920": product_img_urls
}
new_df = pd.DataFrame(data)
print(new_df)
new_df.to_csv("images_to_odoo.csv", encoding='utf-8', index=False)
