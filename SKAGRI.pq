let
    Source = (Parameter1 as binary) => let
        Source = Excel.Workbook(Parameter1, null, true),
        #"First Sheet1" = Source{0}[Data],
        #"Removed Top Rows" = Table.Skip(#"First Sheet1",3),
        #"Removed Bottom Rows" = Table.RemoveLastN(#"Removed Top Rows",6),
        #"Promoted Headers" = Table.PromoteHeaders(#"Removed Bottom Rows", [PromoteAllScalars=true]),
        #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"STT", Int64.Type}, {"Ngay Giao dich", type text}, {"Ngay Hach toan", type text}, {"So Du truoc", Int64.Type}, {"+/-", type text}, {"So Tien Giao dich", Int64.Type}, {"So Du", Int64.Type}, {"Loai Tien", type text}, {"Dien Giai", type text}, {"Chi nhanh", Int64.Type}, {"Trang thai", type text}}),
        #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{{"Ngay Hach toan", "Ngày hạch toán"}, {"Dien Giai", "Diễn giải"}}),
        #"Changed Type with Locale" = Table.TransformColumnTypes(#"Renamed Columns", {{"Ngày hạch toán", type date}}, "en-GB"),
        #"Changed Type1" = Table.TransformColumnTypes(#"Changed Type with Locale",{{"Tu tai khoan", type text}, {"Den tai khoan", type text}, {"So Du truoc", type number}, {"So Tien Giao dich", type number}, {"So Du", type number}}),
        #"Added Custom" = Table.AddColumn(#"Changed Type1", "Ngày giao dịch", each [Ngay Giao dich]),
        #"Replaced Value" = Table.ReplaceValue(#"Added Custom","h ",":",Replacer.ReplaceText,{"Ngày giao dịch"}),
        #"Replaced Value1" = Table.ReplaceValue(#"Replaced Value","' ",":",Replacer.ReplaceText,{"Ngày giao dịch"}),
        #"Changed Type with Locale1" = Table.TransformColumnTypes(#"Replaced Value1", {{"Ngày giao dịch", type datetime}}, "en-GB"),
        #"Added Custom1" = Table.AddColumn(#"Changed Type with Locale1", "Số tài khoản NH", each if [#"+/-"] = "+" then [Den tai khoan] else [Tu tai khoan]),
        #"Changed Type2" = Table.TransformColumnTypes(#"Added Custom1",{{"Số tài khoản NH", type text}}),
        #"Added Custom2" = Table.AddColumn(#"Changed Type2", "Phát sinh tăng tiền", each [So Tien Giao dich] * (if [#"+/-"] ="+" then 1 else 0)),
        #"Changed Type3" = Table.TransformColumnTypes(#"Added Custom2",{{"Phát sinh tăng tiền", type number}}),
        #"Added Custom3" = Table.AddColumn(#"Changed Type3", "Phát sinh giảm tiền", each [So Tien Giao dich] * (if [#"+/-"] = "-" then 1 else 0)),
        #"Changed Type4" = Table.TransformColumnTypes(#"Added Custom3",{{"Phát sinh giảm tiền", type number}}),
        #"Renamed Columns1" = Table.RenameColumns(#"Changed Type4",{{"Ngày hạch toán", "Ngày hạch toán gốc"}}),
        #"Duplicated Column" = Table.DuplicateColumn(#"Renamed Columns1", "Ngày giao dịch", "Ngày giao dịch - Copy"),
        #"Renamed Columns2" = Table.RenameColumns(#"Duplicated Column",{{"Ngày giao dịch - Copy", "Ngày hạch toán"}}),
        #"Changed Type with Locale2" = Table.TransformColumnTypes(#"Renamed Columns2", {{"Ngày hạch toán", type date}}, "en-GB"),
        #"Removed Columns" = Table.RemoveColumns(#"Changed Type with Locale2",{"Ngày hạch toán gốc"}),
        #"Added Index" = Table.AddIndexColumn(#"Removed Columns", "Index", 1, 1, Int64.Type),
        #"Renamed Columns3" = Table.RenameColumns(#"Added Index",{{"Index", "Số chứng từ (STT bank)"}})
    in
        #"Renamed Columns3"
in
    Source