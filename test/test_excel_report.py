import itertools
from typing import Set

import pytest


def pytest_generate_tests(metafunc):
    """
    Parameterize metafunc so that it is run once per model.

    For every run all the fixtures are parameterized with one model name.
    Args:
        metafunc(Metafunc):
            Represents a test function and its decorators, class etc.
            see https://docs.pytest.org/en/4.6.x/reference.html#metafunc

    Returns: None

    """
    argnames = metafunc.fixturenames[:-1]

    all_models = set(metafunc.cls.models)
    model_markers = {
        model: [pytest.mark.model(model)]
        for model in all_models
    }
    for model in get_marked_models(metafunc, "xfail_models"):
        model_markers[model].append(pytest.mark.xfail(strict=True))
    for model in get_marked_models(metafunc, "skip_models"):
        model_markers[model].append(pytest.mark.skip)
    argvalues = [
        pytest.param(
            *[model for _ in range(len(argnames))],
            marks=model_markers[model],
        )
        for model in all_models
    ]
    metafunc.parametrize(
        argnames, argvalues, indirect=True, ids=all_models, scope="class"
    )


def get_marked_models(metafunc, marker) -> Set[str]:
    """
    Collect the content of all pytest.mark.[marker] decorators on metafunc.

    E.g. if marker == "xfail_models", the corresponding decorator is pytest.mark.xfail_models
    Args:
        metafunc(Metafunc):
            Represents a test function and its decorators, class etc.
            see https://docs.pytest.org/en/4.6.x/reference.html#metafunc
        marker(str):
            The name of the Mark

    Returns: set of all model names contained in the decorator pytest.mark.[marker]

    """
    marks = getattr(metafunc.function, "pytestmark", [])
    selected_marks = [list(mark.args) for mark in marks if mark.name == marker]
    selected_models = set(itertools.chain.from_iterable(selected_marks))
    return selected_models


class TestExcelReportA(object):
    models = ["fzkhaus", "noebauer"]

    def test_excel_report_01(self, ifc_file):
        """
        Scenario: test_excel_report_01
        """
        assert True

    @pytest.mark.xfail_models("fzkhaus")
    def test_excel_report_02(self, ifc_file):
        """
        Scenario: test_excel_report_02
        """
        assert True

    @pytest.mark.skip_models("fzkhaus")
    def test_excel_report_03(self, ifc_file):
        """
        Scenario: test_excel_report_01
        """
        assert True

    @pytest.mark.xfail_models("fzkhaus")
    def test_excel_report_04(self, ifc_file):
        """
        Scenario: test_excel_report_05
        """
        assert False

    def test_excel_report_05(self, ifc_file):
        """
        Scenario: test_excel_report_06
        """
        assert True is False
