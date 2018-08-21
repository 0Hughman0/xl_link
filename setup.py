from distutils.core import setup
from setuptools.config import read_configuration

setup(name='xl_link',
    version='0.133dev',
    description='Excel spreadsheets made easy through Pandas magic',
    author='Hugh Ramsden',
    url='https://github.com/0Hughman0/xl_link',
    download_url="https://github.com/0Hughman0/xl_link/archive/0.132.tar.gz",
    packages=['xl_link', 'xl_link.xlsxwriter', 'xl_link.xl_types'],
    license='MIT',
    install_requires='pandas>=0.19',
    extras_require={'xlsxwriter': 'xslxwriter>=0.9',
                    'openpyxl': 'openpyxl>=2.4'},
    python_require=">=3.4")

conf_dict = read_configuration('./setup.cfg')