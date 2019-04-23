from . import metadata, utility
import openpyxl as xlsx
import xlrd as xls
from typing import List, Dict, Iterable
import io

class WorkbookHelper:
    _wb: xls.Book
    _wb_type: str
    def __init__(self, xl_bytes = io.BufferedIOBase):
        try:
            self._wb = xls.open_workbook(file_contents=xl_bytes.read())
            self._wb_type = "xls"
        except:
            self._wb = None
            self._wb_type = "err"

    def get_sheet_data(self, sheet_name: str, range: str) -> Iterable[Dict]:
        for coords in utility.cell_to_coords(range):
            yield {
                "row": coords[0],
                "col": coords[1],
                "cell": f"{utility.number_to_char(coords[1])}{coords[0]}",
                "value": self._wb.sheet_by_name(sheet_name).cell_value(*coords)
            }

"""

class TemplateHelper(WorkbookHelper):
    templates: List[metadata.TemplateMetadata] = []
    def __init__(self, xl_bytes):
        super().__init__(xl_bytes)

    def load_templates(self):
        for s in self._wb.sheetnames:
            ws = self.load_sheet(s)
            t = metadata.TemplateMetadata()
            t.name = s
            t.top_left = (ws.min_column, ws.min_row)
            t.bottom_right = (ws.max_column, ws.max_row)
            self.templates = self.templates + [t]
    
    def load_sheet(self, sheetname: str):
        return self._wb.get_sheet_by_name(sheetname)
"""
