from enum import Enum

class DataClassCard:
    def __init__(self, id, total, rows):
        # id no in first row
        self.id = id
        self.total = total
        self.rows = rows

    def __str__(self):
        # return str(vars(self))
        
        rows = map(lambda row: '{}'.format(', '.join(map(lambda row: f"{row}", row))), self.rows)

        return """
        id: {}
        total: {}
        rows: {}
        """.format(self.id, self.total, '\n\t'.join(rows))



class Loc:
    def __init__(self, min_row, max_row, min_column, max_column):
        self.min_row = min_row
        self.max_row = max_row
        self.min_column = min_column
        self.max_column = max_column


class REVENUE_TYPES(Enum):
    # this related to `first_col` param of write_unit function
    A = '2023'
    B = '2024'

    def is_a(self):
        return self == REVENUE_TYPES.A