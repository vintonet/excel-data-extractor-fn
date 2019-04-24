from . import utility, repo
import openpyxl as xlsx
import xlrd as xls
from typing import List, Dict, Iterable, Union
import io

#todo: add handling for xlsx
class WorkbookHelper:
    _repo: repo.WorkBookRepository

    def __init__(self, xl_bytes = io.BufferedIOBase):
        repos = []
        exceptions: List[Exception] = []
        
        for r in repo.repositories:
            try: 
                r.load_workbook(xl_bytes)
                repos = repos + [r]
            except Exception as e:
                exceptions = exceptions + [e]

        if len(repos) == 0:
            CR = "\r\n"
            raise Exception(f"Could not load workbook:{CR.join([str(e) for e in exceptions])}") from exceptions[0]

        else: 
            _repo = repos[0]


    def get_sheet_data(self, sheet_name: str, range: str) -> Iterable[repo.CellMetadata]:
        return self._repo.get_sheet_data(sheet_name, range)


