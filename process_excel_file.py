import json
from typing import Union, List, Any, Tuple
import string

from openpyxl import Workbook, load_workbook
import win32com.client as win32
import pandas as pd
from pandas import DataFrame
import matplotlib.pyplot as plt
from matplotlib import pyplot


class ExcelProcessor:
    """XlsxWriter can only create new files. It cannot read or modify existing files."""

    def __init__(self, excel_file: str, skip_rows=None, header=0):
        self.excel_file = excel_file
        self.skip_rows = skip_rows
        self.header = header
        if len(pd.ExcelFile(self.excel_file).sheet_names) > 1:
            self.data = self._read_multiple_sheets
        else:
            self.data = self._read_single_sheet()

    @staticmethod
    def jsonify(dict_object: dict):
        return json.dumps(dict_object)

    def _read_single_sheet(self, sheet_name=None) -> DataFrame:
        return pd.read_excel(io=self.excel_file,
                             sheet_name=sheet_name if sheet_name else 0,
                             skiprows=self.skip_rows,
                             header=self.header)

    def select_cols(self, cols: List) -> DataFrame:
        # https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
        return self.data.loc[:, tuple(cols)]

    @property
    def _read_multiple_sheets(self) -> DataFrame:
        all_sheets = []
        excel_df = None
        xlsx = pd.ExcelFile(self.excel_file)
        for sheet in xlsx.sheet_names:
            all_sheets.append(xlsx.parse(sheet,
                                         skiprows=self.skip_rows,
                                         header=self.header))
            excel_df = pd.concat(all_sheets)
        return excel_df

    def get_stats(self):
        print(self.data.shape)
        print(self.data.head())
        print(self.data.describe())

    def get_column_headers(self) -> List:
        return self.data.columns.tolist()

    @staticmethod
    def write_multiple_dfs_to_worksheets(file_name: str,
                                         df_sheet_names: dict,
                                         start_row=0, start_col=0,
                                         index=False,
                                         header=True):
        # https://xlsxwriter.readthedocs.io/example_pandas_multiple.html
        _writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
        for df, sheet_name in df_sheet_names.items():
            df.to_excel(_writer,
                        sheet_name=sheet_name,
                        startrow=start_row,
                        startcol=start_col,
                        index=index,
                        header=header)
        _writer.save()

    def set_cols(self, cols: List):
        try:
            assert len(cols) == len(self.data.columns), \
                "number of columns passed {}, expected {}".format(len(cols), len(self.data.columns))
            self.data.columns = cols
        except AssertionError as error:
            print(error)

    def new_cols(self, col_name: str, col_vale: Any):
        self.data[col_name] = col_vale

    def sort_data(self, column: Union[str, List[str]], ascending=True):
        return self.data.sort_values(column, axis=0, ascending=ascending)

    @staticmethod
    def set_col_as_category_type(df: DataFrame, col_name: str, categories: List):
        df[col_name] = df[col_name].astype('category')
        df[col_name].cat.set_categories(categories, inplace=True)

    def plot_data(self,
                  sel_cols: str,
                  chart_kind: str,
                  num_of_rows: int = None,
                  figsize: Tuple[int, float] = None) -> pyplot:
        self.data[sel_cols].head(num_of_rows).plot(kind=chart_kind, figsize=figsize)
        return plt.show()

    def col_mean(self, col: str):
        return self.data[col].mean()

    def pivot_table(self, index_cols: Union[str, List[str]],
                    subset_cols: List = None,
                    reset_index=True):
        try:
            subset_data = self.data[subset_cols] if List else self.data
            assert isinstance(index_cols, (str, List)), "value must either be a str or a list."
            if isinstance(index_cols, str):
                if reset_index:
                    return subset_data.pivot_table(index=[index_cols]).reset_index()
                return subset_data.pivot_table(index=[index_cols])
            if isinstance(index_cols, List):
                if reset_index:
                    return subset_data.pivot_table(index=index_cols).reset_index()
                return subset_data.pivot_table(index=index_cols)
        except AssertionError as error:
            print(error)

    def save_to_excel(self,
                      file_name: str,
                      sheet_name: str,
                      index: bool = False):
        with self.writer(file_name) as _writer:
            self.data.to_excel(excel_writer=_writer,
                               index=index,
                               sheet_name=sheet_name)
            writer.save()

    @staticmethod
    def add_sheet(_wb, sheet_name):
        """ worksheet names defaults to Sheet1, Sheet2 etc., override by specifying a name"""
        return _wb.add_workshee(sheet_name)

    @staticmethod
    def write_data_to_sheet(_ws, row, col, _data, cell_format=None):
        """ rows and columns are zero indexed. The first cell in a ws, A1, is (0, 0)."""
        return _ws.write(row, col, _data, cell_format)

    @staticmethod
    def writer(file_name,
               engine: str = 'xlsxwriter',
               date_format: str = 'YYYY-MM-DD',
               datetime_format: str = 'YYYY-MM-DD HH:MM:SS',
               mode: str = 'w'):
        return pd.ExcelWriter(path=file_name, engine=engine, mode=mode,
                              date_format=date_format, datetime_format=datetime_format)

    def get_xls_writer_objects(self, df: DataFrame, file_name, sheet_name):
        """ In order to apply XlsxWriter features such as Charts, Conditional Formatting
        and Column Formatting to the Pandas output we need to access the underlying
        workbook and worksheet objects. After that we can treat them as normal XlsxWriter objects."""
        # https://xlsxwriter.readthedocs.io/working_with_pandas.html
        # https://github.com/webermarcolivier/xlsxpandasformatter

        # Get the xlsxwriter workbook and worksheet objects.
        _writer = self.writer(file_name)  # This is a new Pandas Excel writer

        # Convert the dataframe to an XlsxWriter Excel object.
        df.to_excel(_writer, sheet_name=sheet_name, index=False)

        _workbook = _writer.book
        _worksheet = _writer.sheets[sheet_name]
        return _writer, _workbook, _worksheet

    @staticmethod
    def zoom(_ws, value: int):
        return _ws.set_zoom(value)

    def set_row_format(self, df: DataFrame, file_name, row, sheet_name, cell_format, height=None):

        _writer, _workbook, _worksheet = self.get_xls_writer_objects(df,
                                                                     file_name,
                                                                     sheet_name)

        # Add formats to cell
        cell_format = _workbook.add_format(cell_format)

        _worksheet.set_row(row=row,
                           height=height,
                           cell_format=cell_format)
        _writer.save()
        _writer.close()

    @staticmethod
    def set_column_format(_ws,
                          cell_range: str,
                          cell_format: dict = None,
                          first_col: int = None,
                          last_col: int = None,
                          width=None,
                          options: dict = None):

        if all([first_col, last_col]):
            _ws.set_column(first_col, last_col, width, cell_format, options)
        elif cell_range:
            _ws.set_column(cell_range, width, cell_format, options)

    @staticmethod
    def conditional_format_cell(_ws, cell_range: str, options: dict):
        # first_row, first_col, last_row, last_col
        return _ws.conditional_format(cell_range, options)

    @staticmethod
    def auto_fit_cols(file_name: str, sheet_name: str):
        excel_app = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel_app.Workbooks.Open(file_name)
        ws = wb.Worksheets(sheet_name)
        ws.Columns.AutoFit()
        wb.Save()
        excel_app.Application.Quit()

    @staticmethod
    def max_col_str_len(df):
        def max_col_len():
            """ max len of string value in each column"""
            import numpy as np
            length = np.vectorize(len)
            return length(df.values.astype(str)).max(axis=0).tolist()

        def col_name_len():
            """ len of column names in a dataframe"""
            return [len(x) for x in df.columns]

        return [max(col_len, col_val) for col_len, col_val in zip(max_col_len(), col_name_len())]

    def col_auto_fit(self, df: DataFrame, ws):
        for idx, width in enumerate(self.max_col_str_len(df)):
            ws.set_column(idx, idx, width)

    @staticmethod
    def cell_format(_wb, format_dict: dict):
        return _wb.add_format(format_dict)

    def number_of_rows(self):
        """ the number of rows in order to place the totals"""
        return len(self.data.index)

    @staticmethod
    def add_formula(_ws, cell_location, formula, cell_format):
        return _ws.write_formula(cell_location, formula, cell_format)


if __name__ == '__main__':

    excel = ExcelProcessor('movies.xls')
    data = excel.data
    print(excel.number_of_rows())

    sorted_data = excel.sort_data('Gross Earnings', False)
    writer, workbook, worksheet = excel.get_xls_writer_objects(df=data,
                                                               file_name='output.xlsx',
                                                               sheet_name='test1')
    excel.set_row_format(df=data,
                         file_name='output.xlsx',
                         cell_format={'bg_color': '#FFC7CE',
                                      'bold': True,
                                      'font_color': '#9C0006'},
                         row=0,
                         height=18.75,
                         sheet_name='test1')
    format1 = workbook.add_format({'bg_color': '#C6EFCE',
                                   'font_color': '#006100',
                                   'num_format': '#,###'})

    format2 = workbook.add_format({'bg_color': '#FFC7CE',  # Light red fill with dark red text.
                                   'font_color': '#9C0006',
                                   'num_format': '#,###'})
    format3 = workbook.add_format({'num_format': '#,###',
                                   'align': 'right', })

    for i in ['I', 'J'] + list(string.ascii_uppercase)[14:25]:
        # order of conditional formatting rule matters
        excel.conditional_format_cell(worksheet,
                                      '{0}{1}:{0}{2}'.format(i, worksheet.dim_rowmin + 2,
                                                             worksheet.dim_rowmax),
                                      {'type': 'average',
                                       'criteria': 'above',
                                       'format': format2})

        excel.conditional_format_cell(worksheet,
                                      '{0}{1}:{0}{2}'.format(i, worksheet.dim_rowmin + 2,
                                                             worksheet.dim_rowmax),
                                      {'type': 'cell',
                                       'criteria': '>',
                                       'value': 30000000,
                                       'format': format1})

        excel.set_column_format(worksheet,
                                cell_format=format3,
                                cell_range='{0}{1}:{0}{2}'.format(i, worksheet.dim_rowmin + 2, worksheet.dim_rowmax),
                                width=13.86)
    excel.col_auto_fit(data, worksheet)
    excel.zoom(worksheet, 90)
    workbook.close()

    pvt = excel.pivot_table(index_cols=['Language'], subset_cols=['Language', 'Budget', 'Gross Earnings'])
    pvt.to_excel('pivot.xlsx')
    wtr, wb, ws = excel.get_xls_writer_objects(df=pvt,
                                               file_name='pivot_formatted.xlsx',
                                               sheet_name='sheet1')
    format4 = excel.cell_format(wb, {'num_format': '#,###',
                                     'align': 'right', })

    format5 = excel.cell_format(wb, {'bold': False, })

    for i in ['B', 'C']:
        excel.set_column_format(ws,
                                cell_range='{0}{1}:{0}{2}'.format(i, ws.dim_rowmin + 2,
                                                                  ws.dim_rowmax),
                                cell_format=format4)

    excel.set_column_format(ws,
                            cell_range='{0}{1}:{0}{2}'.format('A', ws.dim_rowmin + 2,
                                                              ws.dim_rowmax),
                            cell_format=format5)
    excel.col_auto_fit(pvt, ws)
    wb.close()
