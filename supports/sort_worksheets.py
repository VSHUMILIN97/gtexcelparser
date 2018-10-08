from xlwt import Workbook


class XLWTWorkbook(Workbook):

    @property
    def worksheets(self):
        return self._Workbook__worksheets

    @worksheets.setter
    def worksheets(self, worksheets):
        self._Workbook__worksheets = worksheets
