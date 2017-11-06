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


class AbstractChartWrapper:

    def __init__(self, xlmap, type_, writer, subtype=None):
        self.xlmap = xlmap
        self.type_ = type_
        self.writer = writer
        self.subtype = subtype

    @abstractmethod
    def add_series(self, name, categories, values):
        pass


class XlsxWriterChartWrapper(AbstractChartWrapper):

    def __init__(self, xlmap, type_, writer, subtype):
        super().__init__(xlmap, type_, writer, subtype)

        self.chart = writer.book.add_chart({'type': type_, 'subtype': subtype} if subtype else {'type': type_})

    def add_series(self, name, values, categories=None):
        values = "=" + values.frange

        if isinstance(name, XLCell):
            name = "=" + name.fcell

        kwargs = {'name': name, 'values': values}

        if categories:
            kwargs['categories'] = "=" + categories.frange

        self.chart.add_series(kwargs)


class OpenPyXLChartWrapper(AbstractChartWrapper):

    pass


def create_chart(xlmap, writer, type_, values, categories, names, subtype=None):
    print(values, names, categories)

    engine = writer.engine
    values = ensure_list(values)
    categories = ensure_list(categories)

    if engine not in imported:
        importlib.import_module(engine)
        imported.append(engine)

    if engine == "xlsxwriter":
        chart = XlsxWriterChartWrapper(xlmap, type_, writer, subtype)

    if len(categories) == 1 and len(categories) < len(values):
        categories *= len(values)

    print(values, names, categories)

    for name, category, value in zip(names, categories, values):
        print(name, category, value)
        chart.add_series(name, value, category)

    return chart.chart