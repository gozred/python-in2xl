
from setuptools import setup
import os

VERSION = '0.2.0'
DESCRIPTION = 'A package that allows to insert data in an XLSX file without changing the layout.'


def read(fname):
    with open(os.path.join(os.path.dirname(__file__), fname), 'r') as file:
        return file.read()


setup(
    name="in2xl",
    version=VERSION,
    url='https://github.com/gozred/python-in2xl',
    license='MIT License',
    author="Herzog(gozred)",
    description=DESCRIPTION,
    long_description_content_type="text/markdown",
    long_description=read('pypi.md'),
    package_dir={"in2xl": "src"},
    packages=['in2xl', 'in2xl.in2xl'],


    install_requires=['openpyxl', 'xlsxwriter', 'ruamel.std.zipfile', 'lxml', 'pandas'],
    keywords=['python', 'xlsx', 'excel', 'dataframe', 'insert in excel', 'template', 'excel template'],
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: Developers",
        "Programming Language :: Python :: 3",
        "Operating System :: Unix",
        "Operating System :: MacOS :: MacOS X",
        "Operating System :: Microsoft :: Windows",
    ]
)
