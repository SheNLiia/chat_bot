import os

import requests
from dotenv import load_dotenv

load_dotenv(".env")

TOKEN = os.getenv("YANDEX_TOKEN")


def get_file():
    headers = {'Authorization': TOKEN}
    link = 'https://cloud-api.yandex.net/v1/disk/resources/download'
    download_link = requests.get(link, params={"path": f"Yandex.Forms/690df7f5068ff0fbd8626059/2025-11-14 КМПО.json"},
                                 headers=headers)
    print(download_link.json())
    file = requests.get(download_link.json()["href"])
    return file


def write_file(file_name, file):
    with open(file_name, 'wb') as f:
        f.write(file.content)


if __name__ == "__main__":
    write_file("new.json", get_file())
