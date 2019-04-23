def cell_to_coords(cell_name: str):
    if ":" in cell_name:
        return range_to_coords(cell_name)
    else:
        col_num = char_to_number(cell_name[0])
        row_num = int(cell_name[1:]) - 1
        return [(row_num,col_num)]


def range_to_coords(cell_range: str):
        (tl_cell_name, br_cell_name) = cell_range.split(":")
        (tl_row_num, tl_col_num) = cell_to_coords(tl_cell_name)[0]
        (br_row_num, br_col_num) = cell_to_coords(br_cell_name)[0]
        for col in range(tl_col_num, br_col_num):
            for row in range(tl_row_num, br_row_num):
                yield (row,col)

def get_range_top_left(range: str) -> (int, int):
    if ":" in range:
        range = range.split(":")[0]
    return cell_to_coords(range)[0]

def char_to_number(char: str) -> int:
    return ord(char.lower()) - 97

def number_to_char(num: int) -> str:
    return chr(num + 97).upper()
    