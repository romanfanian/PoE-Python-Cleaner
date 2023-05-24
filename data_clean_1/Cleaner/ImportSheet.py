from typing import List

from openpyxl.cell import Cell

from Cleaner import ImportTable
from Cleaner import Sheet


class ImportSheet(Sheet):

    def __init__(self, file_name: str):
        super().__init__(file_name)

    """@staticmethod
    def is_empty(data_line: list) -> bool:
        for cell in data_line:
            if cell.value is not None and str(cell.value).strip() != '':
                return False
        return True

    def correct_column_length(self) -> list:
        new_column_line = []
        for column in self.column_list:
            if not self.is_empty(column):
                new_column_line.append(column)
        return new_column_line

    def correct_row_length(self) -> list:
        new_row_line = []
        for row in self.row_list:
            if not self.is_empty(row):
                new_row_line.append(row[:len(self.column_list)])
        return new_row_line

    def get_table_amount(self) -> int:
        count = 0
        for cell in self.column_list[0]:
            if cell.value and (type(cell.value) == str and self.table_number
                               in cell.value) or \
                    cell.fill.start_color.index == self.yellow_fill:
                count += 1
        return count

    def get_block_indices(self) -> list:
        indices = []
        cell_index = 0
        for cell in self.column_list[0]:
            if cell.value and (type(cell.value) == str and self.table_number
                               in cell.value) or \
                    cell.fill.start_color.index == self.yellow_fill:
                indices.append(cell_index + 1)
            cell_index += 1
        indices.append(len(self.row_list) + 1)
        return indices

    def partition_header(self) -> list:
        header_blocks = []
        size_of_header = 3
        for block_index in range(len(self.block_indices)):
            curr = self.block_indices[block_index]
            block_header = []
            length_of_header = curr + size_of_header
            for row in self.row_list[curr:length_of_header]:
                block_header.append(row)
            header_blocks.append(block_header)
        return header_blocks

    def partition_products(self) -> list:
        product_blocks = []
        for block_index in range(len(self.block_indices) - 1):
            curr = self.block_indices[block_index]
            curr_next = self.block_indices[block_index + 1]
            product_blocks.append(self.column_list[0][curr:curr_next - 1])
        return product_blocks

    def partition_numbers(self, length: int) -> list:
        numbers_block = []
        for block_index in range(len(self.block_indices) - 1):
            curr = self.block_indices[block_index] + length
            curr_next = self.block_indices[block_index + 1]
            column_block = []
            for column in self.column_list[1:]:
                column_block.append(column[curr:curr_next - 1])
            numbers_block.append(column_block)
        return numbers_block

    def partition(self) -> None:
        self.column_list = list(self.sheet_source.columns)
        self.row_list = list(self.sheet_source.rows)
        self.column_list = self.correct_column_length()
        self.row_list = self.correct_row_length()
        self.table_amount = self.get_table_amount()
        self.block_indices = self.get_block_indices()  # maybe be changed depending on type of sheet
        self.header_list = self.partition_header()
        self.products_list = self.partition_products()
        self.numbers_list = self.partition_numbers(len(self.header_list) - 1)
        
    def clean(self):
        self.read_files()
        self.partition()
        self.fill_partition()
        for block in self.partitions:
            block.clean()"""

    def build_block(self, index: int, header: List[List[Cell]],
                    products: List[Cell], data: List[List[Cell]]) \
            -> ImportTable:
        """
        Creates a single Block object and returns it using header, products and
        places/numbers (data)
        :param index:
        :param header:
        :param products:
        :param data:
        :return: ImportTable
        """
        table = ImportTable(index, header, products, data,
                            self.places_file_data,
                            self.canadian_places, self.products_file_data,
                            self.number_quantity_file_data)
        return table
