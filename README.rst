in2xl
#########

 Readme.rst is still under construction


About this project
*******************************

In the Python programming language, there exist various proficient tools to write data to XLSX format. Two of the most commonly used tools are `XlsxWriter <https://pypi.org/project/XlsxWriter/>`_ and `openpyxl <https://pypi.org/project/openpyxl>`_. With these tools, one can conveniently create an Excel file or write data to an existing Excel file. However, there are some limitations to be aware of when using these tools. Specifically, `XlsxWriter <https://pypi.org/project/XlsxWriter/>`_ is capable of creating files but not modifying them, whereas `openpyxl <https://pypi.org/project/openpyxl>`_ can modify files but does not retain all formatting.

In the data science domain, such limitations can pose challenges, especially if employees have significantly edited Excel files and only require updated data. To address this issue, in2xl offers a simplistic and efficient solution. This tool enables users to transfer data and data frames directly into an Excel file without affecting the existing formatting. Hence, in2xl can be a useful tool for data scientists seeking to update data in pre-existing Excel files.

Table of Contents
*****************

.. contents:: 
    :depth: 2

Install
*****************

in2xl is available on pypi.org. Simply run ``pip install in2xl`` to install it.

Requirements: >= Python 3.7

Project dependencies installed by pip:
::
    lxml
    pandas
    openpyxl
    ruamel.std.zipfile
    XlsxWriter

Usage
*****

The names of the functions are intentionally adapted to `openpyxl <https://pypi.org/project/openpyxl>`_ to make them easier to use and to adapt existing scripts. 

Open a Workbook
""""""""""""""""

It is not possible to create new workbooks using in2xl. The intended approach is to open an existing Excel file (xlsx), insert data, and save it. The opened file serves as a template, where a copy is generated and modified to suit the requirements.

*Example 1:*

..  code-block:: python

    from in2xl import Workbook
    
    wb = Workbook().load_workbook(path)

But this method is also possible:

*Example 2:*

..  code-block:: python

    import in2xl as ix
    
    wb = ix.load_workbook(path)


Open a Worksheet
""""""""""""""""

..  code-block:: python
 
    ws = wb[sheetname]

Insert data
"""""""""""""


Save & Close
"""""""""""""

..  code-block:: python
 
    ws.save(path)
    ws.close()


Additional functions
"""""""""""""""""""""


Planned further functions
"""""""""""""""""""""


