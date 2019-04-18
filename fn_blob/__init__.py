import logging

import azure.functions as func
import tablib as t
import json
import io

def main(inblob: func.InputStream, outdoc: func.Out[func.Document]):
    logging.info(f"Python blob trigger function processed blob \n"
                 f"Name: {inblob.name}\n"
                 f"Blob Size: {inblob.length} bytes")
    
    #load blob into memory and extract contents as json 
    try:
        xls_bytes = inblob.read()
        data_dict = extract_xls_data(xls_bytes)
        outdoc.set(func.Document.from_dict({ "items": data_dict }))
    except Exception as e:
        logging.exception(str(e))
        raise e

def extract_xls_data(xls_bytes: io.BytesIO):
    xls_dataset = t.Dataset().load(in_stream=xls_bytes)
    #xls_json_list = xls_dataset.json
    #xls_obj_list = json.loads(xls_json_list)
    #xls_json_str = json.dumps(xls_obj_list)
    return xls_dataset.dict
    

