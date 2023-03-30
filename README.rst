in2xl
#########

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
    

Additionally you can check the names of all worksheets

..  code-block:: python

    print(wb.sheetnames)


Insert data
"""""""""""""

Different types of data can be inserted directly via ``insert()``

..  code-block:: python
 
    ws.insert(df,2,3, header=False)
..

More detailed description of this function:

>>> insert(data, row=1, column=1, axis=0, header=True, index=False)


 Parameters: 
   **data:   Union(str, int, float, decimal, pd.DataFrame)**
             Besides strings and real numbers, DataFrames can also be inserted directly.
   **row:    int**
             The row in which the data is to be inserted. The default is the first row.
   **column: int**
             The column in which the data is to be inserted. The default is the first column.
   **axis:   int**
             Specify whether the data is inserted in the original orientation or a transposed direction. Default is 0 
             0 : If the data is in a vertical orientation, it will be inserted vertically. 
             1 : If the data is in a vertical orientation, it will be inserted horizontally.
   **header: bool**
             True to include headers in the data, False otherwise. Default is **True**.
   **index:  bool**
             True to include index in the data, False otherwise. Defaults to **False**.  
             

Save & Close
"""""""""""""

..  code-block:: python
 
    ws.save(path)
    ws.close()

The file can be saved multiple times (under different names). As long as the file has not been closed, the temporary Excel file exists. The close command deletes this temporary file.


Additional functions
"""""""""""""""""""""

Template files are sometimes created for multiple tasks/situations. Not all worksheets are always necessary for this. To be able to use these files anyway, it is possible to hide these worksheets. 

..  code-block:: python
   
   print(wb.wb_state) # Returns the visibility status of all worksheets
   print(ws.state) # Returns the visibility status of the current worksheet
   
   ws.state = 0 # Sets the visibility status to 'visible'.
   ws.state = 1 # Sets the visibility status to 'hidden'. User can make this worksheet visible again out of Excel via "Unhide".
   ws.state = 2 # Sets the visibility status to 'veryHidden'. Worksheet is not visible under "Unhide" in Excel.

Planned further functions
"""""""""""""""""""""

* Insert Data into tables / update range of the tables
* Refresh Data of a pivot table
* delete worksheets


