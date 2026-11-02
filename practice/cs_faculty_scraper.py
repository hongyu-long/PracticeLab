"""
There isn't any robot.txt in "https://cs.arizona.edu/about/faculty"

This file scrapes data from above website, then make it in to excel sheet!
"""

import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin

from io import BytesIO
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image as PILImage

import logging


logging.basicConfig(
    level=logging.INFO,
    format="[%(levelname)s] %(asctime)s : %(message)s",
    )

def request_website(url):
    """
    return: html
    """
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        
    except requests.RequestException as error:
        logging.warning(error)
        return None
    
    logging.info(f"Success! {url}")
    return response.text    # get src code


def get_faculty(url):
    html = request_website(url)

    soup = BeautifulSoup(html, "html.parser")

    ppl = soup.select(".az-person-row") # ppl:list
    logging.info(f"found {len(ppl)} faculties")
    faculties = []

    for person in ppl:
        name = person.select_one(".field--name-title").text
        photo_src = person.select_one(".az-person-row-img img").get("src")

        photo_url = urljoin(url, photo_src)
        email = person.select_one(".field--name-field-az-email a").get_text(strip=True)
        
        faculties.append({
            "name": name,
            "pic": photo_url,
            "email": email
        })
    logging.info("finished!")
    return faculties

def store_to_excel(all_faculty : list):
    book = Workbook()
    sheet = book.active
    sheet.title = "cs_faculty"

    sheet.append(["name", "photo", "email"])
    # sheet.write(0, 3, "info")

    photo_size = 200

    sheet.column_dimensions["A"].width = 25
    sheet.column_dimensions["B"].width = photo_size / 7 + 3
    sheet.column_dimensions["C"].width = 35

    for row, ppl in enumerate(all_faculty, start=2):
        sheet.append([ppl["name"], None, ppl["email"]])
        sheet.row_dimensions[row].height = (photo_size + 10) *0.75
        # sheet.write(row, 3, ppl["info"])
        if not ppl["pic"]:
            continue

        # just b/c it returns webp
        try:
            response = requests.get(ppl["pic"], timeout=10)
            response.raise_for_status()

            with PILImage.open(BytesIO(response.content)) as original:
                photo = original.convert("RGB")
                photo.thumbnail((photo_size, photo_size))
                buffer = BytesIO()

                photo.save(buffer, format="PNG")

            buffer.seek(0)
            sheet.add_image(ExcelImage(buffer), f"B{row}")

        except (requests.RequestException, OSError) as error:
            logging.warning(f"Photo failed from {ppl["name"]} with {error}")

    book.save(u"CSFaculty.xlsx")
    logging.info("Saved to CSFaculty.xlsx")

def main():
    url = "https://cs.arizona.edu/about/faculty"

    store_to_excel(get_faculty(url))

main()
