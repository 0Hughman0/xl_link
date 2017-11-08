from unittest import defaultTestLoader, TestSuite

from . import xlmap, xl_types, indexers


def load_tests(loader, standard_tests, pattern):
    suite = TestSuite()
    frame_proxy_tests = defaultTestLoader.loadTestsFromModule(xlmap)
    xl_types_tests = defaultTestLoader.loadTestsFromModule(xl_types)
    indexer_tests = defaultTestLoader.loadTestsFromModule(indexers)
    suite.addTests(frame_proxy_tests)
    suite.addTests(xl_types_tests)
    suite.addTest(indexer_tests)
    return suite
