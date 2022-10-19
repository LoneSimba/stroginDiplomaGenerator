import json
import os
import unicodedata

import openpyxl
import re
from docx import Document as D
from docx.document import Document
from docx.table import Table
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from src import MailCloudDownloader


# main() start
def main():
    config: dict = load_config()
    comp: str = config.get("comp")
    ingest_path: str = os.path.join(os.getcwd(), "files/", comp)

    # table_name = MailCloudDownloader.download_table(config.get('table_link'), ingest_path)
    table_name = config.get('table_name')

    if table_name is None:
        raise Exception('No table was downloaded')

    table: Workbook = openpyxl.load_workbook(os.path.join(ingest_path, table_name))
    table_conf: dict = config.get('table').get(comp)

    for i in range(0, table_conf.get('cats') - 1):
        ws: Worksheet = table.worksheets[i]
        if not re.findall(r'\d\s', str(ws.title)):
            continue

        nomination: str = re.sub(r'\d\s', '', str(ws.title))
        start_row: int = 6

        for row in ws.iter_rows(min_row=start_row, values_only=True):
            if not row[0]:
                break

            values: list = [row[v] for k, v in table_conf.get('cols').items()]
            entry: dict = dict(zip(table_conf.get('cols').keys(), values))
            entry.update({'nomination': nomination})

            if 'аннулирован' in entry.get('result'):
                continue

            count: int = int(re.findall(r'\d', str(entry.get('count')))[0])
            is_group: bool = count > 1
            is_individual: bool = 'Самостоятельный участник' in entry.get('form')
            is_prized: bool = any(x in entry.get('result') for x in ['Дипломант', 'Лауреат', 'Гран-при'])

            output_path: str = os.path.join(
                os.getcwd(), "output", comp,
                str((entry.get('tutor'), '')[is_individual]).strip()
            )

            in_filename: str = "шаблон_"
            in_filename += ("благодарность", "диплом")[is_prized]
            in_filename += ("_участник", "_участник_группа")[is_group]
            in_filename += ("", "_инд")[is_individual]
            in_filename += ".docx"

            if not is_individual:
                tutor_ent: dict = {'tutor': entry.get('tutor'), 'group': entry.get('group'),
                                   'school': entry.get('school')}
                create_tutor_dipl(tutor_ent, ingest_path, output_path)

            if not is_group:
                create_diploma(entry, os.path.join(ingest_path, in_filename), output_path, is_prized)
                continue

            participants = re.sub(r',', ', ', entry.get('participant'))
            participants = participants.split(',')

            if len(participants) != count:
                create_diploma(entry, os.path.join(ingest_path, in_filename), output_path, is_prized)
                continue

            for partic in participants:
                en = dict(entry)
                en.update({'participant': partic})
                create_diploma(en, os.path.join(ingest_path, in_filename), output_path, is_prized)


# main() end


def create_diploma(entry: dict, ingest_path: str, output_path: str, is_prized: bool):
    out_filename: str = ("благодарность", "диплом")[is_prized]
    out_filename += re.sub(r'[\"\'\\/?%*:|<>]', '',
                           (" " + entry.get('participant') + " " + entry.get('title') + ".docx"))
    output_path = unicodedata.normalize('NFKD', output_path)
    out_filename = unicodedata.normalize('NFKD', os.path.join(output_path, out_filename))

    temp: Document = D(ingest_path)
    keys: list = list(entry.keys())
    table: Table = temp.tables[0]
    for cell in table.column_cells(0):
        par = cell.paragraphs[0]
        for run in par.runs:
            text = str(run.text).replace('%', '')
            if text in keys:
                run.text = entry.get(text) or ''

    os.makedirs(output_path, exist_ok=True)
    temp.save(out_filename)


def create_tutor_dipl(entry: dict, ingest_path: str, output_path: str):
    out_filename: str = "благодарность"
    out_filename += re.sub(r'["\'\\/?%*:|<>]', '',
                           (" " + entry.get('tutor') + " " + entry.get('group') + " " + entry.get('school') + ".docx"))
    output_path = unicodedata.normalize('NFKD', output_path)
    out_filename = unicodedata.normalize('NFKD', os.path.join(output_path, out_filename))

    if os.path.isfile(out_filename):
        return

    temp: Document = D(os.path.join(ingest_path, "шаблон_благодарность_педагог.docx"))
    keys: list = list(entry.keys())
    table: Table = temp.tables[0]
    for c in table.column_cells(0):
        p = c.paragraphs[0]
        for r in p.runs:
            text = str(r.text).replace('%', '')
            if text in keys:
                r.text = entry.get(text) or ''
    os.makedirs(output_path, exist_ok=True)
    temp.save(os.path.join(output_path, out_filename))


# load_config() start
def load_config() -> dict:
    try:
        with open('conf.json', 'r', encoding='utf-8') as conf:
            config: dict = json.load(conf)
    except OSError:
        print("Failed to open conf.json file")
        exit(1)
    except json.JSONDecodeError:
        print("conf.json contents are broken")
        exit(1)

    return config


# load_config() end


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print("Fatal error: " + str(e))
