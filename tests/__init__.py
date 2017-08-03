from unittest import defaultTestLoader, TestSuite

from . import frame_proxy, xl_types


def load_tests(loader, standard_tests, pattern):
    suite = TestSuite()
    frame_proxy_tests = defaultTestLoader.loadTestsFromModule(frame_proxy)
    xl_types_tests = defaultTestLoader.loadTestsFromModule(xl_types)
    suite.addTests(frame_proxy_tests)
    suite.addTests(xl_types_tests)
    return suite
