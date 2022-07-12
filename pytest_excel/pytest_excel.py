import re
from datetime import datetime
from collections import OrderedDict
from openpyxl import Workbook
import pytest
from _pytest.mark.structures import Mark

_py_ext_re = re.compile(r"\.py$")


def pytest_addoption(parser):
    group = parser.getgroup("terminal reporting")
    group.addoption(
        "--excelreport",
        "--excel-report",
        action="store",
        dest="excelpath",
        metavar="path",
        default=None,
        help="create excel report file at given path.",
    )


def pytest_configure(config):
    excelpath = config.option.excelpath
    if excelpath:
        config._excel = ExcelReporter(excelpath)
        config.pluginmanager.register(config._excel)


def pytest_unconfigure(config):
    excel = getattr(config, "_excel", None)
    if excel:
        del config._excel
        config.pluginmanager.unregister(excel)


def mangle_test_address(address):
    path, possible_open_bracket, params = address.partition("[")
    names = path.split("::")
    try:
        names.remove("()")
    except ValueError:
        pass

    names[0] = names[0].replace("/", ".")
    names[0] = _py_ext_re.sub("", names[0])
    names[-1] += possible_open_bracket + params
    return names


class ExcelReporter(object):
    def __init__(self, excelpath):
        self.results = []
        self.row_key = "model"
        self.column_key = "test_step"
        self.wbook = Workbook()
        self.wsheet = None
        self.rc = 1
        self.excelpath = datetime.now().strftime(excelpath)

    def append(self, result):
        self.results.append(result)

    def create_sheet(self, column_heading):

        self.wsheet = self.wbook.create_sheet(index=0)

        all_row_fields = sorted(set([data[self.row_key] for data in self.results]))
        for i, row_label in enumerate(all_row_fields, 2):
            self.wsheet.cell(row=i, column=1).value = row_label

        all_col_fields = sorted(set([data[self.column_key] for data in self.results]))
        for i, col_label in enumerate(all_col_fields, 2):
            self.wsheet.cell(row=1, column=i).value = col_label

        # for heading in column_heading:
        #     index_value = column_heading.index(heading) + 1
        #     heading = heading.replace("_", " ").upper()
        #     self.wsheet.cell(row=self.rc, column=index_value).value = heading
        # self.rc = self.rc + 1

    def update_worksheet(self):
        all_row_fields = sorted(set([data[self.row_key] for data in self.results]))
        all_col_fields = sorted(set([data[self.column_key] for data in self.results]))
        for data in self.results:
            col_idx = all_col_fields.index(data[self.column_key]) + 2
            row_idx = all_row_fields.index(data[self.row_key]) + 2
            value = "OK" if data["result"] == "PASSED" else "ERR"
            try:
                self.wsheet.cell(row=row_idx, column=col_idx).value = value
            except ValueError:
                pass

        # for data in self.results:
        #     for key, value in data.items():
        #         try:
        #             self.wsheet.cell(row=self.rc, column=list(data).index(key) + 1).value = value
        #         except ValueError:
        #             self.wsheet.cell(row=self.rc, column=list(data).index(key) + 1).value = str(vars(value))
        #     self.rc = self.rc + 1

    def save_excel(self):
        self.wbook.save(filename=self.excelpath)

    def build_result(self, report, status, message):

        result = OrderedDict()
        names = mangle_test_address(report.nodeid)

        model_name = None
        for mark in report.test_marker:
            if mark.name == "model":
                model_name = mark.args[0]
                break
        if model_name is None:
            raise AttributeError(
                f"Test {report.nodeid} is not marked with a model name.\n"
                "Add @pytest.mark.model('my_model_name') to the test function."
            )
        result["model"] = model_name

        # test name is something like 'test_excel_report_01[fzkhaus]'
        test_name = names[-1]
        step_name_parts = test_name.split("[")[0].split("_")[1:]
        result["test_step"] = "_".join(step_name_parts)

        if report.test_doc is None:
            result["description"] = report.test_doc
        else:
            result["description"] = report.test_doc.strip()

        result["result"] = status
        result["duration"] = getattr(report, "duration", 0.0)
        result["timestamp"] = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
        result["message"] = message
        result["file_name"] = report.location[0]

        self.append(result)

    def append_pass(self, report):
        status = "PASSED"
        message = None
        self.build_result(report, status, message)

    def append_failure(self, report):

        if hasattr(report, "wasxfail"):
            status = "XPASSED"
            message = "xfail-marked test passes Reason: %s " % report.wasxfail

        else:
            if hasattr(report.longrepr, "reprcrash"):
                message = report.longrepr.reprcrash.message
            elif isinstance(report.longrepr, str):
                message = report.longrepr
            else:
                message = str(report.longrepr)

            status = "FAILED"

        self.build_result(report, status, message)

    def append_error(self, report):

        message = report.longrepr
        status = "ERROR"
        self.build_result(report, status, message)

    def append_skipped(self, report):

        if hasattr(report, "wasxfail"):
            status = "XFAILED"
            message = "expected test failure Reason: %s " % report.wasxfail

        else:
            status = "SKIPPED"
            _, _, message = report.longrepr
            if message.startswith("Skipped: "):
                message = message[9:]

        self.build_result(report, status, message)

    def build_tests(self, item):

        result = OrderedDict()
        names = mangle_test_address(item.nodeid)

        result["suite_name"] = names[-2]
        result["test_name"] = names[-1]
        if item.obj.__doc__ is None:
            result["description"] = item.obj.__doc__
        else:
            result["description"] = item.obj.__doc__.strip()
        result["file_name"] = item.location[0]
        test_marker = []
        test_message = []
        for k, v in item.keywords.items():
            if isinstance(v, list):
                for x in v:
                    if isinstance(x, Mark):
                        if x.name != "usefixtures":
                            test_marker.append(x.name)
                        if x.kwargs:
                            test_message.append(x.kwargs.get("reason"))

        test_markers = ", ".join(test_marker)
        result["markers"] = test_markers

        test_messages = ", ".join(test_message)
        result["message"] = test_messages
        self.append(result)

    def append_tests(self, item):

        self.build_tests(item)

    @pytest.mark.trylast
    def pytest_collection_modifyitems(self, session, config, items):
        """called after collection has been performed, may filter or re-order
        the items in-place."""
        if session.config.option.collectonly:
            for item in items:
                self.append_tests(item)

    @pytest.mark.hookwrapper
    def pytest_runtest_makereport(self, item, call):

        outcome = yield

        report = outcome.get_result()
        report.test_doc = item.obj.__doc__
        test_marker = []
        for k, v in item.keywords.items():
            if isinstance(v, list):
                for x in v:
                    if isinstance(x, Mark):
                        test_marker.append(x)
            if isinstance(v, Mark):
                test_marker.append(v)
        report.test_marker = test_marker

    def pytest_runtest_logreport(self, report):

        if report.passed:
            if report.when == "call":  # ignore setup/teardown
                self.append_pass(report)

        elif report.failed:
            if report.when == "call":
                self.append_failure(report)

            else:
                self.append_error(report)

        elif report.skipped:
            self.append_skipped(report)

    def pytest_sessionfinish(self, session):
        if not hasattr(session.config, "slaveinput"):
            if self.results:
                fieldnames = list(self.results[0])
                self.create_sheet(fieldnames)
                self.update_worksheet()
                self.save_excel()

    def pytest_terminal_summary(self, terminalreporter):
        if self.results:
            terminalreporter.write_sep("-", "excel report: %s" % self.excelpath)
