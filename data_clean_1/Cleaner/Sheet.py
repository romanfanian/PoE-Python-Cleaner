from typing import List
import openpyxl.utils.exceptions
from openpyxl import load_workbook
from openpyxl.cell import Cell
from Cleaner.Table import Table


class Sheet(Table):
    """
    Sheet class for any type of Excel sheet
    """
    places_file = '../csv/places'
    products_file = '../csv/products'
    number_quantity_file = '../csv/number_quantity'
    table_headers = ['ARTICLES EXPORTED', 'ARTICLES',
                     'EXPORTED', 'IMPORTED']
    yellow_fill = 'FFFFFF00'

    def __init__(self, file_name: str):
        self.file_name = file_name
        self.sheet_object = None
        self.sheet_source = None
        self.id_assoc = False
        self.partitions = []
        self.places_file_data = []
        self.canadian_places = []
        self.products_file_data = []
        self.number_quantity_file_data = []
        self.row_list = []
        self.column_list = []

    def get_places_data(self):
        return self.places_file_data

    def get_canadian_places_data(self):
        return self.canadian_places

    def get_products_data(self):
        return self.products_file_data

    def get_numbers_data(self):
        return self.number_quantity_file_data

    def get_partitions(self):
        return self.partitions

    def get_sheet_source(self):
        return self.sheet_source

    def get_block_distances(self) -> List[int]:
        """
        Gets the distances between each block, prints if something
        sussy is going on
        :return:
        """
        dist = []
        for table_index in range(1, len(self.partitions)):
            prev_table = self.partitions[table_index - 1]
            curr_table = self.partitions[table_index]
            prev_index = prev_table.get_start_index()
            curr_index = curr_table.get_start_index()
            print(prev_index)
            distance = curr_index - prev_index
            dist.append(distance)
            """if not 40 <= distance <= 70:
                print("LOOK AT " + str(curr_index))"""
        return dist

    @staticmethod
    def is_data_line_empty(data_line: List[Cell]) -> bool:
        """
        Determines if a line of data is empty
        :param data_line:
        :return: boolean
        """
        for cell in data_line:
            cell_value = cell.value
            if cell_value is not None and str(cell_value).strip() != '':
                return False
        return True

    def insert_columns(self, number: int) -> None:
        """
        Inserts number amount of empty columns to the sheet
        :param number:
        :return:
        """
        pass


    def find_column_start(self) -> int:
        """
        Function finds where the list of columns can start
        :return:
        """
        for column_index, column in enumerate(self.column_list):
            if not self.is_data_line_empty(column):
                return column_index
        return 0

    def correct_column_dimension(self) -> None:
        """
        Modifies list of columns by excluding empty columns
        :return:
        """
        new_column_list = []
        for column in self.column_list:
            if not self.is_data_line_empty(column):
                new_column_list.append(column)
        self.column_list = new_column_list

    def correct_dimensions(self) -> None:
        """
        Modifies list of rows by excluding empty columns
        :return:
        """
        column_start = self.find_column_start()
        self.correct_column_dimension()
        for row_index, row in enumerate(self.row_list):
            column_length = len(self.column_list)
            self.row_list[row_index] = row[column_start:column_length]

        #TODO
        #  Correct ID Association here!!
        #  It should restrict the row amount by some number

    @staticmethod
    def empty_cell(cell: Cell):
        """
        Checks if a given Cell object is empty
        :param cell:
        :return:
        """
        if cell.value is None:
            return True
        if isinstance(cell.value, str) and \
                cell.value.strip() == '':
            return True
        return False

    def build_block(self, index: int, header: List[List[Cell]],
                    products: List[Cell], data: List[List[Cell]]) -> Table:
        """
        Creates a single Block object and returns it using header, products and
        places/numbers (data)
        :param index:
        :param header:
        :param products:
        :param data:
        :return: Block
        """
        return Table(index, header, products, data, self.places_file_data,
                      self.canadian_places, self.products_file_data,
                      self.number_quantity_file_data)

    def partition(self) -> None:
        """
        Breaks the entire Sheet objects data into Block objects of data
        containing each table's information
        :return:
        """
        self.correct_dimensions()
        products_column = self.column_list[0]
        for cell_index, cell in enumerate(products_column):
            cell_value = cell.value
            if self.is_table_header(cell_value):
                self.correct_block_index(cell_index)
                header = self.partition_header(cell_index)
                products = self.partition_product(cell_index, len(header))
                data = self.partition_data(cell_index, len(header))
                block = self.build_block(cell_index, header, products, data)
                self.partitions.append(block)

    def correct_block_index(self, index: int) -> int:
        """
        Access the self.partitions and enable Block classes to hold their own
        block header index, use this paired with correct_block_indices()
        :param index:
        :return int:
        """
        prev_block_start_index = self.get_prev_block_start()
        prev_row = self.row_list[prev_block_start_index][1:]
        while not self.is_data_line_empty(prev_row) and index >= 0:
            index -= 1
        return index

    @staticmethod
    def is_empty_cell(cell: Cell):
        """
        Determines if a cell is empty or contains no information
        :param cell:
        :return:
        """
        cell_value = cell.value
        if cell_value is None:
            return True
        if isinstance(cell_value, str) and \
                cell_value.strip() == '':
            return True
        return False

    @staticmethod
    def contains_number(data: List[Cell]) -> bool:
        """
        Determines if there exists at least one number in a given row/column
        :param data:
        :return: boolean
        """
        for cell in data:
            cell_value = cell.value
            cell_value_type = type(cell_value)
            if cell_value and (cell_value_type == int or
                               (cell_value_type == str and
                                cell_value.strip().isdigit())):
                return True
        return False

    def get_prev_block_start(self) -> int:
        """
        Returns previous Block's start index
        :return:
        """
        if self.partitions:
            block = self.partitions[-1]
            return block.get_start_index()
        return 0

    def get_next_block_start(self, index: int) -> int:
        """
        Returns the next Block's start index
        :param index:
        :return:
        """
        products = self.column_list[0]
        step = 1
        next_block_start = index + step
        next_cell = products[next_block_start]
        while not self.is_table_header(next_cell.value) \
                and next_block_start < len(products) - 1:
            step += 1
            next_block_start = index + step
            next_cell = products[next_block_start]
        return next_block_start

    def partition_header(self, index: int) -> List[List[Cell]]:
        """
        Function grabs header information about a Block using its header index
        :param index:
        :return: List[List[Cell]]
        """
        size_of_header = 0
        block_header = []
        length_of_header = index + size_of_header
        rows = self.row_list[index:]
        row = rows[size_of_header]
        while size_of_header < len(rows) - 1 and not \
                self.contains_number(row):
            size_of_header += 1
            row = rows[size_of_header]
            length_of_header = index + size_of_header
        for row in self.row_list[index:length_of_header]:
            block_header.append(row)
        return block_header

    def partition_product(self, index: int, size_of_header: int) -> List[Cell]:
        """
        Function grabs product information from the Sheet using a header index
        and the next Table's starting index
        :param size_of_header:
        :param index:
        :return:
        """
        products = self.column_list[0]
        next_block_start = self.get_next_block_start(index)
        end_of_header = index + size_of_header
        return products[end_of_header:next_block_start - 1]

    def partition_data(self, index: int, size_of_header: int) -> \
            List[List[Cell]]:
        """
        Function grabs data (places/numbers) information from the sheet using
        the starting index of a Table.
        :param size_of_header:
        :param index:
        :return:
        """
        next_block_start = self.get_next_block_start(index)
        column_block = []
        for column in self.column_list[1:]:
            start_index = index + size_of_header
            end_index = next_block_start - 1
            column_block.append(column[start_index:end_index])
        return column_block

    def clean(self) -> None:
        """
        Cleans all the Table's in a Sheet
        :return:
        """
        for table in self.partitions:
            table.clean()

    def id_assoc_check(self) -> bool:
        """
        Detects whether a sheet has run the ID Association tool on it
        :return:
        """
        pass

    def open_files(self) -> None:
        """
        Opens all the files necessary and grabs information
        :return:
        """
        try:
            self.sheet_object = load_workbook(self.file_name)
            self.sheet_source = self.sheet_object.active
            self.column_list = list(self.sheet_source.columns)
            self.row_list = list(self.sheet_source.rows)
            self.id_assoc = self.id_assoc_check()
        except openpyxl.utils.exceptions.InvalidFileException:
            pass
        with open(self.places_file, encoding='utf-8') as file:
            for line in file:
                self.places_file_data.append(line.strip('\n').split(','))
        with open(self.products_file, encoding='utf-8') as file:
            for line in file:
                self.products_file_data.append(line.strip('\n'))
        with open(self.number_quantity_file, encoding='utf-8') as file:
            for line in file:
                self.number_quantity_file_data.append(line.strip('\n'))
        for place in self.places_file_data:
            self.canadian_places.append(place[0])

    def save(self, filename: str) -> None:
        """
        Saves sheet to a new file
        :param filename:
        :return:
        """
        self.sheet_object.save(filename)
