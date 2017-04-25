import matplotlib.pyplot as plt
from matplotlib import colors
import pandas as pd
import numpy as np
import re
from xlsxwriter.utility import xl_range, xl_rowcol_to_cell



def convert_colormap_to_hex(cmap, x, vmin=0, vmax=1):
    """
    Example::
        >>> seaborn.palplot(seaborn.color_palette("RdBu_r", 7))
        >>> colorMapRGB = seaborn.color_palette("RdBu_r", 61)
        >>> colormap = seaborn.blend_palette(colorMapRGB, as_cmap=True, input='rgb')
        >>> [convert_colormap_to_hex(colormap, x, vmin=-2, vmax=2) for x in range(-2, 3)]
        ['#09386d', '#72b1d3', '#f7f6f5', '#e7866a', '#730421']
    """
    norm = colors.Normalize(vmin, vmax)
    color_rgb = plt.cm.get_cmap(cmap)(norm(x))
    color_hex = colors.rgb2hex(color_rgb)
    return color_hex


class FormatedWorksheet:
    """
    A FormatedWorksheet is a helper class that wraps the worksheet, workbook and dataframe objects
    written by pandas to_excel method using the xlsxwriter engine. The FormatedWorksheet class
    takes care of keeping record of cells format as a table of dictionaries and allowing user to update format of columns, rows
    and cells.

    See REAMDE.mk for a detailed example.
    """

    def __init__(self, worksheet, workbook, df, hasIndex, *args, **kwargs):
        self.worksheet = worksheet
        self.workbook = workbook
        self.df = df
        self.nColLevels = df.columns.nlevels
        self.nIndexLevels = df.index.nlevels
        self.hasIndex = hasIndex

        if self.hasIndex:
            self.nIndexCol = self.nIndexLevels
        else:
            self.nIndexCol = 0

        self.nHeaderRow = self.nColLevels
        if self.hasIndex and self.nIndexLevels > 1:
            # One additional row for multiindex names
            self.nHeaderRow += 1

        self.nRows = len(df)
        self.nCols = len(df.columns)

        # We want to format also multiindex columns, so we convert the non-multiindex columns to tuples
        self.dfColumns = [None for x in range(self.nCols)]
        for iCol in range(self.nCols):
            if type(self.df.columns[iCol]) is not tuple:
                self.dfColumns[iCol] = (self.df.columns[iCol], )
            else:
                self.dfColumns[iCol] = self.df.columns[iCol]

        self.formatTable = [[{} for j in range(self.nCols)] for i in range(self.nRows)]
        # We use the same format for all header cells at a specific level
        self.formatTableHeader = [{} for i in range(self.nColLevels)]
        self.headerRowsHeight = [None for i in range(self.nColLevels)]
        # We use the same format for all index cells at a specific level
        self.formatTableIndex = [{} for i in range(self.nIndexLevels)]
        self.indexColWidth = [None for j in range(self.nIndexLevels)]


    def apply_format_table(self):
        
        # Apply format to dataframe cells
        for index, row in self.df.iterrows():

            rowIndex = self.df.index.get_loc(index)
            iRow, worksheetRow = self.convert_to_row_index(index)

            for iCol in range(self.nCols):

                iCol, worksheetCol = self.convert_to_col_index(iCol)
                x = row.iloc[iCol]
                cell = xl_rowcol_to_cell(worksheetRow, worksheetCol)
                formatDic = self.formatTable[rowIndex][iCol]
                cellFormat = self.workbook.add_format(formatDic)
                self.worksheet.write(cell, x, cellFormat)

        # Apply format to header cells
        for i in range(self.nColLevels):
            formatDic = self.formatTableHeader[i]
            cellFormat = self.workbook.add_format(formatDic)
            self.worksheet.set_row(i, self.headerRowsHeight[i], cellFormat)

        # Apply format to index cells
        for j in range(self.nIndexLevels):
            formatDic = self.formatTableIndex[i]
            cellFormat = self.workbook.add_format(formatDic)
            self.worksheet.set_column(xl_range(1, j, 1, j), self.indexColWidth[j], cellFormat)


    def convert_to_col_index(self, col):

        if self.nColLevels > 1:
            expectedColType = tuple
        else:
            expectedColType = str

        if type(col) is expectedColType:
            iCol = self.df.columns.get_loc(col)
        elif type(col) is int:
            iCol = col

        worksheetCol = iCol + self.nIndexCol

        return iCol, worksheetCol


    # def convert_to_row_index(self, row):

    #     if self.nIndexLevels > 1:
    #         expectedRowType = tuple
    #     else:
    #         expectedRowType = str

    #     if type(row) is expectedRowType:
    #         iRow = self.df.index.get_loc(row)
    #     elif type(row) is int:
    #         iRow = row

    #     worksheetRow = iRow + self.nHeaderRow

    #     return iRow, worksheetRow


    def convert_to_row_index(self, row):

        if type(row) is int:
            iRow = row
        else:
            iRow = self.df.index.get_loc(row)

        worksheetRow = iRow + self.nHeaderRow

        return iRow, worksheetRow


    def format_col(self, col, colWidth=None, colFormat=None):

        iCol, worksheetCol = self.convert_to_col_index(col)

        if colWidth is not None:
            self.worksheet.set_column(xl_range(1, worksheetCol, 1, worksheetCol), colWidth)
        if colFormat is not None:
            for rowIndex in range(self.nRows):
                self.formatTable[rowIndex][iCol].update(colFormat)


    def format_row(self, row, rowHeight=None, rowFormat=None):

        iRow, worksheetRow = self.convert_to_row_index(row)

        if rowHeight is not None:
            self.worksheet.set_row(worksheetRow, rowHeight)
        if rowFormat is not None:
            for iCol in range(self.nCols):
                self.formatTable[iRow][iCol].update(rowFormat)


    def format_cols(self, colWidthList=None, colFormatList=None, colPatternFormatList=None):
        """
        colFormatList should be a list of dictionary-like options.
        colPatternFormatList should a a list of tuples (pattern for column name, dictionary of format options).
        """
        
        if type(colWidthList) is list and len(colWidthList) != self.nCols:
            print("Warning: length of colWidthList is different from the nb of columns of the dataframe.")
            return
        if type(colFormatList) is list and len(colFormatList) != self.nCols:
            print("Warning: length of colFormatList is different from the nb of columns of the dataframe.")
            return
       
        if type(colWidthList) is list:
            for iCol in range(self.nCols):
                colWidth = colWidthList[iCol]
                self.format_col(iCol, colWidth=colWidth)

        if type(colFormatList) is list:
            for iCol in range(self.nCols):
                colFormat = colFormatList[iCol]
                self.format_col(iCol, colFormat=colFormat)

        if type(colPatternFormatList) is list:
            for iCol in range(self.nCols):
                # Apply the format if we find the pattern at any of the multiindex levels of columns
                for colPattern, formatDic in colPatternFormatList:
                    if np.any([re.search(colPattern, col) for col in self.dfColumns[iCol]]):
                        self.format_col(iCol, colFormat=formatDic)


    def format_numeric_cols(self, colPatternFormatList):
        """
        Formats numeric columns.

        Example::
            numFormatScientific = 0x0b
            numFormatFloat2digits = 0x02
            numFormatInteger = 0x01
            colPatternFormatList = [
                (r'(pvalue)', numFormatScientific),
                (r'(proportion)', '0.0E+00'),
                (r'([oO]ddsRatio)|(odds ratio)', numFormatFloat2digits),
                (r'(cterm bias n seq)|(count)', numFormatInteger)
            ]
        """

        for iCol in range(self.nCols):
            # Apply the format if we find the pattern at any of the multiindex levels
            for colPattern, formatNum in colPatternFormatList:
                if np.any([re.search(colPattern, col) for col in self.dfColumns[iCol]]):
                    for rowIndex in range(self.nRows):
                            self.formatTable[rowIndex][iCol]['num_format'] = formatNum


    def format_background_colormap(self, col, colormap, vmin, vmax):

        iCol, worksheetCol = self.convert_to_col_index(col)

        for index, row in self.df.iterrows():
            x = row.iloc[iCol]
            if pd.notnull(x):
                colorHex = convert_colormap_to_hex(colormap, x, vmin=vmin, vmax=vmax)
                rowIndex = self.df.index.get_loc(index)
                self.formatTable[rowIndex][iCol]['bg_color'] = colorHex


    def format_add_separation_border_between_groups(self, groupCol, borderStyle=2):

        # Finding last rows of grouped dataframe on a multiindex column
        colDf = pd.DataFrame(self.df[groupCol].rename()).reset_index(drop=True).reset_index()
        lastDf = colDf.groupby(by=0).last()
        dfLastInGroupIndexList = lastDf['index'].tolist()
        
        for iRow in dfLastInGroupIndexList:
            for iCol in range(self.nCols):
                self.formatTable[iRow][iCol]['bottom'] = borderStyle


    def format_header(self, headerFormat=None, rowHeight=None):

        if type(headerFormat) is list:
            if len(headerFormat) != self.nColLevels:
                print("ERROR: header format list is not same length as number of column levels.")
            else:
                for i, iFormat in enumerate(headerFormat):
                    self.formatTableHeader[i].update(iFormat)
        elif type(headerFormat) is dict:
            for i in range(self.nColLevels):
                self.formatTableHeader[i].update(headerFormat)

        if type(rowHeight) is list:
            if len(rowHeight) != self.nColLevels:
                print("ERROR: header rowHeight list is not same length as number of column levels.")
            else:
                for i, height in enumerate(rowHeight):
                    self.headerRowsHeight[i] = height
        elif type(rowHeight) is int:
            height = rowHeight
            for i in range(self.nColLevels):
                self.headerRowsHeight[i] = height


    def format_index(self, indexFormat=None, colWidth=None):

        if type(indexFormat) is list:
            if len(indexFormat) != self.nIndexLevels:
                print("ERROR: index format list is not same length as number of index levels.")
            else:
                for i, iFormat in enumerate(indexFormat):
                    self.formatTableIndex[i].update(iFormat)
        elif type(indexFormat) is dict:
            for i in range(self.nIndexLevels):
                self.formatTableIndex[i].update(indexFormat)

        if type(colWidth) is list:
            if len(colWidth) != self.nIndexLevels:
                print("ERROR: index colWidth list is not same length as number of index levels.")
            else:
                for i, width in enumerate(colWidth):
                    self.indexColWidth[i] = width
        elif type(colWidth) is int:
            width = colWidth
            for i in range(self.nIndexLevels):
                self.indexColWidth[i] = width


    def format_pandas_map(self, func, col):
        """
        Applies a conditional formatting to cells, using Pandas map method on dataframe column.

        The function must return a format dictionary.
        """

        iCol, worksheetCol = self.convert_to_col_index(col)

        formatSeries = self.df[col].map(func, axis=1)

        for index, formatDic in formatSeries.iteritems():
            iRow = self.df.index.get_loc(index)
            self.formatTable[iRow][iCol].update(formatDic)


    def format_pandas_apply(self, func, axis=1):
        """
        Applies a conditional formatting to cells, using Pandas apply method on dataframe
        rows (axis=1) / columns (axis=0).

        The function must take a Pandas Series and return a Pandas Series of format dictionary of the same
        size as the dataframe columns / rows.
        """
        formatDf = self.df.apply(func, axis=axis)
        if len(formatDf.columns) != len(self.df.columns):
            print("ERROR")
        elif len(formatDf) != len(self.df):
            print("ERROR")

        for i in range(len(formatDf)):
            for j in range(len(formatDf.columns)):
                self.formatTable[i][j].update(formatDf.iloc[i, j])


    def freeze_header(self):
        self.worksheet.freeze_panes(self.nHeaderRow, 0)


    def freeze_index(self):
        self.worksheet.freeze_panes(0, self.nIndexCol)

    def freeze_index_and_header(self):
        self.worksheet.freeze_panes(self.nHeaderRow, self.nIndexCol)
