from Cleaner import Table


class ImportTable(Table):
    def __init__(self, start_index, header, products, data, places_file,
                 canadian_places, products_file, number_file):
        super().__init__(start_index, header, products, data, places_file,
                         canadian_places, products_file, number_file)

    def clean(self):
        self.clean_places(self.data[0])
        self.clean_numbers(self.number_col)
        self.clean_cts(self.cts_col)
