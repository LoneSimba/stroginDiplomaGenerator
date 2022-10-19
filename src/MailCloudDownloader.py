import os
import re
import shutil
from typing import Union

import requests


def download_table(table_link: str, dest_path: str) -> Union[str, None]:
    weblink = re.findall(r'/public/(\w+/\w+)', table_link)[0]

    items_req = requests.get("https://cloud.mail.ru/api/v4/public/list?weblink=" + weblink)
    links_req = requests.get("https://cloud.mail.ru/api/v2/dispatcher", headers={"referer": table_link})

    if items_req.status_code != 200 | links_req.status_code != 200:
        return None

    items = items_req.json()
    links = links_req.json()

    weblink_get = links.get('body').get('weblink_get')[0].get('url')

    if items.get('type') == "folder":
        item_list = items.get('list')
    else:
        item_list = [items]

    if len(item_list) == 0:
        return None

    item = item_list[0]
    item_link = weblink_get + "/" + item.get('weblink')
    filename = re.sub(r'[\"?><:\\/|*]', '', item.get("name"))

    file = requests.get(item_link, stream=True)

    if file.status_code == 200:
        file.raw.decode_content = True

        with open(os.path.join(dest_path, filename), 'w+b') as f:
            shutil.copyfileobj(file.raw, f)
            f.flush()

    else:
        return None

    return filename
