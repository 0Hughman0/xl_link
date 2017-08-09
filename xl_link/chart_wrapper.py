from xlsxwriter import chart


DEFAULT_CHART_SUBTYPES = {'area': 'percent_stacked',
                          'bar': 'percent_stacked',
                          'column': 'percent_stacked',
                          'scatter': 'straight_with_markers',
                          'radar': 'with_markers'}

class SeriesWrapper:

    def __init__(self, categories, values):

        self.values = values
        self.categories = categories
        self.name = None

        self.line = None
        self.border = None
        self.fill = None

        self.pattern = None
        self.gradient = None

        self.marker = None

        self.trendline = None
        self.smooth = False

        self.x_error_bars = None
        self.y_error_bars = None

        self.data_labels = None

        self.points = None

        self.invert_if_neg = False

        self.overlap = 0
        self.gap = 150


    def add_error_bars(self, x_plus=None, x_minus=None, y_plus=None, y_minus=None):



class ChartWrapperBase:

    def __init__(self, workbook, chart_type, subtype):
        if subtype is None:
            subtype = DEFAULT_CHART_SUBTYPES[chart_type]

        if chart_type in DEFAULT_CHART_SUBTYPES:
            chart = workbook.add_chart({'type': chart_type, 'subtype': subtype})
        else:
            chart = workbook.add_chart({'type': chart_type})

        self.chart = chart

    def add_series(self, categories_list, values_list, x_error_bars_list=None, y_error_bars_list=None):


        x_error_bars_dict_list = []
        y_error_bars_dict_list = []

        if x_error_bars_list:
            for x_error_bars in x_error_bars_list
                plus_series, minus_series = x_error_bars
                error_bars['x_error_bars'] = {'type': 'custom',
                                              'plus_values': plus_series,
                                              'minus_values': minus_series}
                x_error_bars_dict_list.append()

        if y_error_bars_list:
            for y_error_bars in y_error_bars_list:
                plus_series, minus_series = y_error_bars
                error_bars['y_error_bars'] = {'type': 'custom',
                                              'plus_values': plus_series,
                                              'minus_values': minus_series}

        for categories, values in zip(categories_list, values_list):



            self.chart.add_series({'categories': categories,
                                       'values': values} + error_bars)


    def

def create_chart_obj(workbook, chart_type, subtype=None):




def create_