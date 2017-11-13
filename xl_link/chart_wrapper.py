from abc import abstractmethod
import importlib
from xl_link.xl_types import to_series
from warnings import warn

from distutils.version import StrictVersion

imported = {}

s = to_series

SINGLE_CATEGORY_CHARTS = ['bar', 'area', 'doughnut', 'line', 'pie', 'radar', 'stock', 'column']
CATEGORIES_REQUIRED_CHARTS = ['scatter']

MIN_VERSIONS = {'xlsxwriter': '0.9',
                'openpyxl': '2.4'}


def check_engine_compatible(engine):
    min_version = MIN_VERSIONS.get(engine.__name__, 0)
    if StrictVersion(engine.__version__) < StrictVersion(min_version):
        raise ImportError(("excel writer engine not met version requirements for xl_link:"
                           "{} < {}").format(engine.__version__ < min_version))


def ensure_list(specifier):
    """
    if specifier isn't a list or tuple, makes specifier into a list containing just specifier
    """
    if not isinstance(specifier, list) and not isinstance(specifier, tuple):
        return [specifier,]

    return specifier


class AbstractChartWrapper:
    """
    Wraps around excel writer module, for use by create_chart.
    """

    def __init__(self, book, type_, subtype=None, engine=None):
        """
        Abstract: should initialise Chart Wrapper class. Should initialise chart object, and set as self.chart.

        Parameters
        ----------
        book : object
            pd.ExcelWriter.book object typical represents whole excel file, and needed to creat new chart
        type_ : str
            type of chart desired
        subtype : str
            subtype of chart desired
        engine : package
            corresponding to engine used (usefull for fetching objects from that package)
        """
        self.type_ = type_
        self.book = book
        self.subtype = subtype
        self.engine = engine
        self.chart = None

    @abstractmethod
    def add_series(self, name, values, categories=None):
        """
        Abstract: should add a series to chart.

        Parameters
        ----------
        name : str
            representing the name of the series
        values : XLRange
            represending values of series
        categories : XLRange
            representing values of categories
        """
        pass

    @property
    @abstractmethod
    def x_axis_name(self):
        pass

    @x_axis_name.setter
    @abstractmethod
    def x_axis_name(self, value):
        pass

    @property
    @abstractmethod
    def y_axis_name(self):
        pass

    @y_axis_name.setter
    @abstractmethod
    def y_axis_name(self, value):
        pass

    @property
    @abstractmethod
    def title(self):
        pass

    @title.setter
    @abstractmethod
    def title(self, value):
        pass


class XlsxWriterChartWrapper(AbstractChartWrapper):

    def __init__(self, book, type_, subtype):
        super().__init__(book, type_, subtype)

        self.chart = book.add_chart({'type': type_, 'subtype': subtype} if subtype else {'type': type_})

    def add_series(self, name, values, categories=None):
        values = s(values.frange)

        kwargs = {'name': name, 'values': values}

        if categories:
            kwargs['categories'] = s(categories.frange)

        self.chart.add_series(kwargs)

    @property
    def x_axis_name(self):
        return self.chart.x_axis['name']

    @x_axis_name.setter
    def x_axis_name(self, value):
        self.chart.x_axis['name'] = value

    @property
    def y_axis_name(self):
        return self.chart.y_axis['name']

    @y_axis_name.setter
    def y_axis_name(self, value):
        self.chart.y_axis['name'] = value

    @property
    def title(self):
        return self.chart.title_name

    @title.setter
    def title(self, value):
        self.chart.title_name = value


def type_to_openpyxl_chart_name(type_):
    return type_.title() + 'Chart'


class OpenPyXLChartWrapper(AbstractChartWrapper):

    def __init__(self, book, type_, subtype, openpyxl):
        super().__init__(book, type_, subtype, openpyxl)
        self.chart = getattr(openpyxl.chart, type_to_openpyxl_chart_name(type_))()
        if subtype:
            self.chart.type = subtype

    def add_series(self, name, values, categories=None):
        Series = self.engine.chart.series_factory.SeriesFactory

        if categories is None:
            series = Series(values.frange, title=name)
            self.chart.append(series)
        elif self.type_ in SINGLE_CATEGORY_CHARTS:
            if values.is_col:
                values.start.row += -1 # To include top cell as name
            else: # is_row
                warn("Appears you're using rows as values with openpyxl, support for this is a little temperamental, you have been warned!")
                values.start.col += -1
            self.chart.add_data(values.frange, titles_from_data=True)
            self.chart.set_categories(categories.frange)
        else:
            series = Series(values.frange, xvalues=categories.frange, title=name)
            self.chart.append(series)

    @property
    def x_axis_name(self):
        return self.chart.x_axis.title

    @x_axis_name.setter
    def x_axis_name(self, value):
        self.chart.x_axis.title = value

    @property
    def y_axis_name(self):
        return self.chart.y_axis.title

    @y_axis_name.setter
    def y_axis_name(self, value):
        self.chart.y_axis.title = value

    @property
    def title(self):
        return self.chart.title

    @title.setter
    def title(self, value):
        self.chart.title = value

def create_chart(workbook, engine, type_, values, categories, names, subtype=None,
                 title=None, x_axis_name=None, y_axis_name=None):
    """
    Create a chart object corresponding to given engine, within workbook.

    Parameters
    ----------
    workbook : object
        to insert chart into, either XlsxWriter.Workbooks or openpyxl Workbooks.
    engine : str
        representing engine to use, either XlsxWriter.Workbooks or openpyxl Workbooks.
    type_ : str
        representing chart type, see engine docs for options e.g. 'line'.
    values : XLRange or sequence of XLRanges
        use as values for each series/ data set (Excel equivalent to y data).
    categories : XLRange or sequence of XLRanges
        use as categories for each series/ data set (Excel equivalent to x data).
    names : str/ sequence of str
        use as name for each series/ data set (Excel uses this to label each dataset).
    subtype : str
        representing chart subtype, see engine docs for details.

    Returns
    -------
    chart : object
        populated chart object corresponding to engine's chart type.
    """
    if 'openpyxl' in engine:
        engine = 'openpyxl' # Cuz pandas appends version to engine name

    values = ensure_list(values)
    categories = ensure_list(categories)
    names = ensure_list(names)

    if engine not in imported:
        engine_mod = importlib.import_module(engine)
        check_engine_compatible(engine_mod)
        imported[engine] = engine_mod

    if engine == 'xlsxwriter':
        chart = XlsxWriterChartWrapper(workbook, type_, subtype)
    elif engine == 'openpyxl':
        chart = OpenPyXLChartWrapper(workbook, type_, subtype, imported[engine])
    else:
        raise TypeError("Couldn't find chart wrapper for {}".format(engine))

    if len(categories) == 1 and len(categories) < len(values):
        categories *= len(values)

    for name, category, value in zip(names, categories, values):
        chart.add_series(name, value, category)

    if title:
        chart.title = title
    if x_axis_name:
        chart.x_axis_name = x_axis_name
    if y_axis_name:
        chart.y_axis_name = y_axis_name

    return chart.chart