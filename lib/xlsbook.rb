#coding: utf-8
require "win32ole"

# --------------------------------------
# Excelブッククラス
# --------------------------------------
class XlsBook
  def initialize()
    @xls = WIN32OLE.new("Excel.Application")
    @books = nil
  end

  def open(xlsfile, xlssheet)
    @books = @xls.workbooks.open(xlsfile)
    sheets = @books.worksheets(xlssheet)
    sheets
  end

  def load_book(xlsfile, xlssheet, &bk)

    sh = open(xlsfile, xlssheet)

    yield sh

    @xls.displayalerts = false
    # saveasで別の名前で保存をするとエラーになる。
    # オリジナルを別名にコピーしてからそれを編集して上書きするようにすること。
    @books.save
    @xls.displayalerts = true
    @xls.quit
  end
end
