import openpyxl as xl
from typing import List

class CellMetadata:
    cell_type: str
    relative_location: (int, int)
    value: str
    style: object

    def __init__(self, relative_location: (int,int), style: object):
        self.relative_location = relative_location
        self.style = style

class FactMetadata(CellMetadata):
    def __init__(self, **kwargs):
        self.cell_type = "fact"
        super().__init__(**kwargs)

class DimensionMetadata(CellMetadata):
    fact: FactMetadata
    def __init__(self, fact: FactMetadata, **kwargs):
        self.cell_type = "dim"
        self.fact = fact
        super().__init__(**kwargs)
        
class AttributeMetadata(CellMetadata):
    dim: DimensionMetadata
    def __init__(self, dim: DimensionMetadata, **kwargs):
        self.cell_type = "attr"
        self.dim = dim
        super().__init__(**kwargs)

class ValueMetadata(CellMetadata):
    attr: AttributeMetadata
    def __init__(self, attr: AttributeMetadata, **kwargs):
        self.cell_type = "val"
        self.attr = attr
        super().__init__(**kwargs)

class SheetMetadata:
    facts: List[FactMetadata]
    dimensions: List[DimensionMetadata]
    attributes: List[DimensionMetadata]
    values: List[ValueMetadata]
    def __init__(self, name:str):
        self.name = name
        self.facts = []
        self.dimensions = []
        self.attributes = []
        self.values = []

class TemplateMetadata:
    name: str
    facts: List[FactMetadata]
    dimensions: List[DimensionMetadata]
    attributes: List[DimensionMetadata]
    top_left: (int, int)
    bottom_right: (int, int)
