from xl_link import __version__
from distutils.core import setup
from setuptools.config import read_configuration

setup(name='xl_link',
    version=__version__,
    description='Excel spreadsheets made easy through Pandas magic',
    author='Hugh Ramsden',
    url='https://github.com/0Hughman0/xl_link',
    download_url="https://github.com/0Hughman0/xl_link/archive/{}.tar.gz".format(__version__),
    packages=['xl_link', 'xl_link.xlsxwriter', 'xl_link.xl_types'],
    license='MIT',
    install_requires='pandas>=0.19',
    extras_require={'xlsxwriter': 'xslxwriter>=0.9',
                    'openpyxl': 'openpyxl>=2.4'},
    python_requires=">=3.4")

conf_dict = read_configuration('./setup.cfg')
