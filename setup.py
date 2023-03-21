from setuptools import setup, find_packages

VERSION = '0.1.0'
DESCRIPTION = 'Insert String, Integer, Float or Dataframe into an XLSX file'
LONG_DESCRIPTION = 'A package that allows to insert data in an XLSX file without changing the layout.'

# Setting up
setup(
    name="in2xl",
    version=VERSION,
    author="David Herzog",
    description=DESCRIPTION,
    long_description_content_type="text/markdown",
    long_description=LONG_DESCRIPTION,
    packages=find_packages(),
    install_requires=['openpyxl', 'xlsxwriter', 'ruamel.std.zipfile', 'lxml', 'pandas'],
    keywords=['python', 'xlsx', 'excel', 'dataframe', 'insert in excel', 'template', 'excel template'],
    classifiers=[
        "Development Status :: 1 - Planning",
        "Intended Audience :: Developers",
        "Programming Language :: Python :: 3",
        "Operating System :: Unix",
        "Operating System :: MacOS :: MacOS X",
        "Operating System :: Microsoft :: Windows",
    ]
)

