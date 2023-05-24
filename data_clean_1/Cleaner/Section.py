from typing import List
from openpyxl.cell import Cell


class Section:
    """
    Section Object
    """

    def __init__(self, products: List[Cell], row_data_list: List[List[Cell]],
                 column_data_list: List[List[Cell]]) -> None:
        self.row_list = row_data_list
        self.column_list = column_data_list
        self.products = products

    def calculate_sums(self) -> None:
        """
        A function that replaces the cells in the last item of
        row_list with their appropriate sums

        TIP: To write a sum in Excel from code, we must create the string
        of the command. (Excel computes the command on its own)

        Ex: '=SUM(A17:A25)'

        To access a cell we use cell.value and to set a cell's value we use
        cell.value as well for example 'cell.value = 3'. Look into openpyxl's
        documentation for more commands.

        TIP 2: You can get a cell's coordinate by doing cell.coordinate which
        accesses the Cell object's class attribute

        Ex: cell.coordinate -> 'A17'

        Use as many helpers as you think you need. Check test_Section for an example test Good luck!
        :return:
        """
        number_cols = self.column_list[1:]
        for col in number_cols:
            first_cell = col[0]
            last_cell = col[len(col) - 2]
            sum_cell = col[len(col) - 1]
            cmd = "=SUM(" + str(first_cell.coordinate) + ":" + \
                  str(last_cell.coordinate) + ")"
            sum_cell.value = cmd
