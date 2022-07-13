import pytest


class TestExcelReportB(object):

    @pytest.mark.model("my-model")
    def test_excel_report_01(self):
        """
        Scenario: test_excel_report_01
        """
        assert True

    @pytest.mark.xfail(reason="passed Simply", strict=True)
    @pytest.mark.model("my-model")
    def test_excel_report_02(self):
        """
        Scenario: test_excel_report_02
        """
        assert True

    @pytest.mark.skip(reason="Skip for No Reason")
    @pytest.mark.model("my-model")
    def test_excel_report_03(self):
        """
        Scenario: test_excel_report_01
        """
        assert True

    @pytest.mark.xfail(reason="Failed Simply")
    @pytest.mark.model("my-model")
    def test_excel_report_04(self):
        """
        Scenario: test_excel_report_05
        """
        assert False

    @pytest.mark.model("my-model")
    def test_excel_report_05(self):
        """
        Scenario: test_excel_report_06
        """
        assert True is False
