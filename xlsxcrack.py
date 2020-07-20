#!/usr/bin/env python3

import argparse
from zipfile import ZipFile
import re


parser = argparse.ArgumentParser(description="An xlsx protection password remover.")
parser.add_argument('filename', help="xlsx file to remove password from")
args = parser.parse_args()

filename = args.filename

with ZipFile(filename, 'r') as original:
    print("zipfile opened...looking for worksheets")
    new_filename = filename.replace('.xlsx', '.cracked.xlsx')
    with ZipFile(new_filename, 'w') as modified:
        for item in original.infolist():
            if re.search(r'xl/worksheets/.*xml$', item.filename):
                print(f"Sheet found {item.filename}...", end='')
                file_data = original.read(item.filename)
                # have to use bytes
                removed = re.sub(rb'<sheetProtection.*?/>', b'', file_data)
                if removed:
                    print('protection removed.')
                    modified.writestr(item, removed)
                else:
                    print('no protection found.')
            # otherwise copy the file to new archive as is
            else:
                file_data = original.read(item.filename)
                modified.writestr(item, file_data)