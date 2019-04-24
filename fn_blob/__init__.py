import logging
import json
import io
import azure.functions as func
import datetime
from . import helpers, utility
from typing import Dict

#updated 240419 2003

def main(inblob: func.InputStream, outdoc: func.Out[func.Document]):
    logging.info(f"Python blob trigger function processed blob \n"
                 f"Name: {inblob.name}\n"
                 f"Blob Size: {inblob.length} bytes")
    
    #load blob into memory and extract contents as json 
    try:
        #todo: refactor as http trigger where you can specify in route
        #currently unable to owing to corrupted inblob when written as trigger for unknown reason
        data_wb = load_workbook(inblob._io, inblob.name.split(".")[1])
        extraction_cfg = load_cfg('./fn_blob/gas.json')
        sheets_processed = []
        sheets_not_found = []
        to_process = extraction_cfg["template_applications"]
        sheet_names = data_wb.get_sheet_names()
        data_catalog = {}

        for item in to_process:
            sheet_name = item["sheet"]
            if sheet_name not in [s for s in sheet_names]:
                sheets_not_found = sheets_not_found + [sheet_name]
            else:
                if sheet_name not in data_catalog.keys():
                    data_catalog[sheet_name] = {}
                template_name = item["template"]
                template = get_template(extraction_cfg, template_name)
                fact = template["fact"]
                fact_data_catalog = {}

                for extractor in template["extractors"]:
                    for dim in extractor["dim"]:
                        if dim not in fact_data_catalog.keys():
                            fact_data_catalog[dim] = {}
                            for attr in extractor["attr"]:
                                if attr not in fact_data_catalog[dim].keys():
                                    fact_data_catalog[dim][attr] = ""

                    extractor_data = data_wb.get_sheet_data(sheet_name, extractor["range"])
                    (tl_row_num, tl_col_num) = utility.get_range_top_left(extractor["range"])

                    for data in extractor_data:
                        dim = extractor["dim"][data["row"]-tl_row_num]
                        attr = extractor["attr"][data["col"]-tl_col_num]
                        fact_data_catalog[dim][attr] = data["value"]

                data_catalog[sheet_name][fact] = fact_data_catalog
                sheets_processed = sheets_processed + [sheet_name]

        data_catalog["metadata"] = {
            "workbook_name": inblob.name,
            "datetime_processed": f"{datetime.datetime.now():%Y-%m-%d %H:%M:%S}",
            "result": "Success",
            "processed": " ,".join(sheets_processed),
            "sheets": " ,".join(sheet_names)
        }

        if len(sheets_not_found) > 0:
            data_catalog["metadata"]["err"] = f"Sheets not found: {' ,'.join(sheets_not_found)}"
            data_catalog["metadata"]["result"] = "Partial Success"

        outdoc.set(func.Document.from_dict({ "items": data_catalog }))

    except Exception as e:
        err = f"Exception processing {inblob.name}: {str(e)}"
        logging.exception(err)
        data_catalog = {}  
        data_catalog["metadata"] = {
            "workbook_name": inblob.name,
            "datetime_processed": f"{datetime.datetime.now():%Y-%m-%d %H:%M:%S}",
            "result": "Failure",
            "error": err
        }
        outdoc.set(func.Document.from_dict({ "items": data_catalog }))

def load_workbook(xls_bytes: io.BufferedIOBase, extension: str) -> helpers.WorkbookHelper:
    return helpers.WorkbookHelper(xls_bytes, extension)

def load_cfg(path: str) -> Dict: 
    with open(path, "r") as f:
        return json.loads(f.read())


def get_template(cfg: Dict, template_name: str) -> Dict:
    try:
        return [t for t in cfg["templates"] if t["template"] == template_name][0]
    except:
        return {}
         