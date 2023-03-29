in2xl
#########

 Readme.rst is still under construction


About this project
*******************************

In the Python programming language, there exist various proficient tools to write data to XLSX format. Two of the most commonly used tools are XlsxWriter and Openpyxl. With these tools, one can conveniently create an Excel file or write data to an existing Excel file. However, there are some limitations to be aware of when using these tools. Specifically, XlsxWriter is capable of creating files but not modifying them, whereas Openpyxl can modify files but does not retain all formatting.

In the data science domain, such limitations can pose challenges, especially if employees have significantly edited Excel files and only require updated data. To address this issue, in2xl offers a simplistic and efficient solution. This tool enables users to transfer data and data frames directly into an Excel file without affecting the existing formatting. Hence, in2xl can be a useful tool for data scientists seeking to update data in pre-existing Excel files.

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
  xlsxwriter
::

