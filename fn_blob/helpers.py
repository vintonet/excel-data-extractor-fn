from . import utility, repo
import openpyxl as xlsx
import xlrd as xls
from typing import List, Dict, Iterable, Union
import io

#todo: add handling for xlsx
class WorkbookHelper:
    _repos: List[repo.WorkBookRepository]

    def __init__(self, xl_bytes: io.BufferedIOBase, extension: str):
        repos = []
        exceptions: List[Exception] = []
        
        for r in repo.repositories:
            try: 
                r.load_workbook(xl_bytes, extension)
                repos = repos + [r]
            except Exception as e:
                exceptions = exceptions + [e]

        if len(repos) == 0:
            CR = "\r\n"
            raise Exception(f"Could not load workbook:{CR.join([str(e) for e in exceptions])}") from exceptions[0]
        else: 
            self._repos = repos

    def get_sheet_data(self, sheet_name: str, range: str) -> Iterable:
        results = None
        exceptions = []
        for r in self._repos:
            try:
                results = r.get_sheet_data(sheet_name, range)    
                break
            except Exception as e:
                exceptions = exceptions + [e]
                pass
            
        if results is not None:
            return results

        else:
            raise Exception(f"could not fetch data due to errors: {', '.join(str(exceptions))}")


    def get_sheet_names(self) -> List[str]:
        results = None
        exceptions = []
        for r in self._repos:
            try:
                results = r.get_sheet_names()
                break
            except Exception as e:
                exceptions = exceptions + [e]
                pass
            
        if results is not None:
            return results

        else:
            raise Exception(f"could not fetch sheet names due to errors: {', '.join(str(exceptions))}")

