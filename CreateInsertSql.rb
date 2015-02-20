# coding: utf-8

require 'win32ole'

# VARCHARの場合、Doubleクォーテーションを付けます。
def _convert_param(type, cell)

    val = cell.value
    text = cell.Text

    if val == nil || val == "null" || val == "NULL" then
        return "null"
    end

    # データ型に応じた変換処理
    if type == "VARCHAR" then
        return "\'#{text}\'"
    elsif type == "DATETIME" then

        if val.instance_of?(String) then
            return "\'#{val}\'"
        else
            return "\'#{val.strftime("%Y-%m-%d %H:%M:%S")}\'"
        end
    else
        return text
    end
end

# 引数で与えられたExcel(パス)を開き、INSERT文を追記します。
def _set_insert_query(filePath)

    tableName = ""
    maxColumnIndex=0
    cullentRow=0
    columnNameList = []
    columnTypeList = []
    dataList = []

    # Excelファイル読み込み
    excel = WIN32OLE.new('Excel.Application')

    begin
        fso = WIN32OLE.new('Scripting.FileSystemObject')
        book = excel.Workbooks.Open(fso.GetAbsolutePathName("#{filePath}"))

        # シートの読み込み(ここでは、一番左のシートしか読み込まない)
        sheet = book.Worksheets(1)
        #puts "SheetName:#{sheet.Name}"

        sheet.UsedRange.Rows.each do |row|
            row.Columns.each do |cell|
                # 列名の取得
                cullentRow = cell.row

                # テーブル名取得
                if cullentRow == 1 && cell.column == 1 then
                    tableName = cell.value
                    puts "TableName:#{tableName}"

                # 列名リスト取得
                elsif cullentRow == 3 && 1 < cell.column then
                    maxColumnIndex = cell.column
                    columnNameList << cell.value

                # データ型リスト取得
                elsif cullentRow == 4 && 1 < cell.column then
                    columnTypeList << cell.value

                # データ取得
                elsif 9 <= cullentRow && 2 <= cell.column then

                    # B列が空のデータ行は処理しない。
                    if 2 == cell.column && cell.value == nil then
                        break
                    end

                    dataList << _convert_param(columnTypeList[cell.column - 2], cell)
                end
            end

            if !dataList.empty? then
                # 空の列の後に、Insert文を設定
                insertQuery = "INSERT INTO ITHost.dbo.#{tableName} \(#{columnNameList.join(",")},INSUSR,INSTNMT,INS_DT,UPDUSR,UPDTNMT,UPD_DT\) VALUES \(#{dataList.join(",")},\'system\',\'system\',GETDATE(),\'system\',\'system\',GETDATE()\);"
                sheet.Cells(cullentRow, maxColumnIndex + 2).value = insertQuery
                #puts insertQuery

                dataList = []
            end
        end
        book.Save
        book.close

    ensure
        excel.quit
    end

#    puts columnNameList.join(",")
#    puts columnTypeList.join(",")
#    puts maxColumnIndex
end

# Excelファイルの一覧を取得して、INSERT文を作成します。
puts "Insert文作成処理を開始します。"
begin
Dir::glob("./**/*.xlsx").each{|filePath|
    puts "処理対象ファイル：#{filePath}"
    _set_insert_query filePath
}
rescue => exp
    puts "処理中に例外が発生しました。"
    p exp
ensure
    puts "Insert文作成処理が完了しました。"
end
