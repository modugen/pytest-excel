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

    xfailed_models = get_xfailed_models(metafunc)
    all_models = set(metafunc.cls.models)
    argvalues = [
        pytest.param(
            *[model for _ in range(len(argnames))],
            marks=pytest.mark.xfail(strict=True) if model in xfailed_models else (),
        )
        for model in all_models
    ]
    metafunc.parametrize(
        argnames, argvalues, indirect=True, ids=all_models, scope="class"
    )


def get_xfailed_models(metafunc) -> Set[str]:
    """
    Collect the content of all pytest.mark.xfail_model decorators on metafunc.
    Args:
        metafunc(Metafunc):
            Represents a test function and its decorators, class etc.
            see https://docs.pytest.org/en/4.6.x/reference.html#metafunc

    Returns: set of all xfailed model names

    """
    marks = getattr(metafunc.function, "pytestmark", [])
    xfail_marks = [list(mark.args) for mark in marks if mark.name == "xfail_models"]
    xfailed_models = set(itertools.chain.from_iterable(xfail_marks))
    return xfailed_models



class Test_Excel_Report_A(object):
  models = ["fzkhaus", "noebauer"]
  
    
  def test_excel_report_01(self, ifc_file):
    """
    Scenario: test_excel_report_01
    """
    assert True
    
  
  @pytest.mark.xfail(reason="passed Simply")  
  def test_excel_report_02(self, ifc_file):
    """
    Scenario: test_excel_report_02
    """
    assert True


  @pytest.mark.skip(reason="Skip for No Reason")
  def test_excel_report_03(self, ifc_file):
    """
    Scenario: test_excel_report_01
    """
    assert True


  @pytest.mark.xfail(reason="Failed Simply")
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



class Test_Excel_Report_B(object):
  
    
  def test_excel_report_01(self):
    """
    Scenario: test_excel_report_01
    """
    assert True
    
  
  @pytest.mark.xfail(reason="passed Simply")  
  def test_excel_report_02(self):
    """
    Scenario: test_excel_report_02
    """
    assert True


  @pytest.mark.skip(reason="Skip for No Reason")
  def test_excel_report_03(self):
    """
    Scenario: test_excel_report_01
    """
    assert True


  @pytest.mark.xfail(reason="Failed Simply")
  def test_excel_report_04(self):
    """
    Scenario: test_excel_report_05
    """
    assert False


  def test_excel_report_05(self):
    """
    Scenario: test_excel_report_06
    """
    assert True is False

