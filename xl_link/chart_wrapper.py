from abc import abstractmethod
import importlib
import xlsxwriter
from xl_link.xl_types import XLCell

imported = []


def ensure_list(specifier):
    """
    if specifier isn't a list or tuple, makes specifier into a list containing just specifier
    """
    if not isinstance(specifier, list) or not isinstance(specifier, tuple):
        return [specifier,]

    return specifier

def type_to_openpyxl_chart_name(type):
    return type.title() + 'Chart'


class AbstractChartWrapper:

    def __init__(self, book, type_, subtype=None):
        self.type_ = type_
        self.sheet = book
        self.subtype = subtype

    @abstractmethod
    def add_series(self, name, categories, values):
        pass


class XlsxWriterChartWrapper(AbstractChartWrapper):

    def __init__(self, book, type_, subtype):
        super().__init__(book, type_, subtype)

        self.chart = book.add_chart({'type': type_, 'subtype': subtype} if subtype else {'type': type_})

    def add_series(self, name, values, categories=None):
        values = "=" + values.frange

        if isinstance(name, XLCell):
            name = "=" + name.fcell

        kwargs = {'name': name, 'values': values}

        if categories:
            kwargs['categories'] = "=" + categories.frange

        self.chart.add_series(kwargs)


class OpenPyXLChartWrapper(AbstractChartWrapper):

    def __init__(self):
        raise NotImplementedError

def create_chart(workbook, engine, type_, values, categories, names, subtype=None):
    """
    Create a chart object corresponding to given engine, within sheet.

    Parameters
    ----------
    workbook : workbook to insert chart into, currently only support XlsxWriter.Workbooks.
    engine : str representing engine to use, currently only supports 'xlsxwriter'.
    type_ : str representing chart type, see engine docs for options e.g. 'line'.
    values : XLRange or sequence of XLRanges to use as values for each series/ data set (Excel equivalent to y data).
    categories : XLRange or sequence of XLRanges, or str/ sequence of strings to use as categories for
        each series/ data set (Excel equivalent to x data).
    names : XLRange or sequence of XLRanges, or str/ sequence of strings to use as name for
        each series/ data set (Excel uses this to label each dataset).
    subtype : str representing chart subtype, see engine docs for details.

    Returns
    -------
    chart object, corresponding to engine's chart type.
    """
    print(values, names, categories)

    values = ensure_list(values)
    categories = ensure_list(categories)

    if engine not in imported:
        importlib.import_module(engine)
        imported.append(engine)

    if engine == "xlsxwriter":
        chart = XlsxWriterChartWrapper(workbook, type_, subtype)

    if len(categories) == 1 and len(categories) < len(values):
        categories *= len(values)

    print(values, names, categories)

    for name, category, value in zip(names, categories, values):
        print(name, category, value)
        chart.add_series(name, value, category)

    return chart.chart