from xlwt import Workbook


class XLWTWorkbook(Workbook):
    """
    Mocking class for extending base realisation and adding a counter for the
    xlwt workbook sheets
    """
    @property
    def worksheets(self):
        return self._Workbook__worksheets

    @worksheets.setter
    def worksheets(self, worksheets):
        self._Workbook__worksheets = worksheets
