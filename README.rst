pytest-excel
================


pytest-excel is a plugin for `py.test <http://pytest.org>`_ that allows to 
to create excel report for test results.

The test results are arranged in a matrix where the columns are the test
names and the rows are the model names.
Passed and xpassed tests are represented as "OK", all other results (failed,
xfailed, skipped) are represented as "ERR".


Requirements
------------

You will need the following prerequisites in order to use pytest-excel:

- Python 3.6, 3.7 and above
- pytest 2.9.0 or newer
- opepyxl
- pandas


Installation
------------

To install pytest-excel::

    $ pip install pytest-excel


Mark the tests you want to run with the name of your model(s)::

    @pytest.mark.model("my-model-name")

Then run your tests with::

    $ py.test --excelreport=report.xls

If you would like more detailed output (one test per line), then you may use the verbose option::

    $ py.test --verbose

If you would like to run tests without execution to collect test doc string::

    $ py.test --excelreport=report.xls --collect-only


If you would like to get timestamp in the as filename::

    $ py.test --excelreport=report%Y-%M-dT%H%.xls
