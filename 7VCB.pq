let
    Source = (Parameter7 as binary) => let
        Source = Excel.Workbook(Parameter7, null, true),
        #"Vietcombank_Account_Statement (_Sheet" = Source{0}[Data],
        #"Renamed Columns" = Table.RenameColumns(#"Vietcombank_Account_Statement (_Sheet",{{"Column1", "Ngày giao dịch"}, {"Column2", "Số tham chiếu"}, {"Column3", "Phát sinh giảm tiền"}, {"Column4", "Phát sinh tăng tiền"}, {"Column5", "Diễn giải"}}),
        #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns",{{"Ngày giao dịch", type text}}),
        #"Thêm cột Ngày đầu năm" = Table.AddColumn(#"Changed Type", "Ngày đầu năm", each if Text.Start([Ngày giao dịch], 9)="Từ: 01/01" then "01-01-"&Text.End(Text.Start([Ngày giao dịch], 14), 4) else null),
        #"Filled Down" = Table.FillDown(#"Thêm cột Ngày đầu năm",{"Ngày đầu năm"}),
        #"Thêm cột ngày hạch toán" = Table.AddColumn(#"Filled Down", "Ngày hạch toán", each if [Ngày giao dịch] ="Ngày giao dịch" then [Ngày đầu năm] else [Ngày giao dịch]),
        #"Added Custom" = Table.AddColumn(#"Thêm cột ngày hạch toán", "Số tài khoản NH", each if [Ngày giao dịch]= "Số tài khoản:" then [Số tham chiếu] else null),
        #"Filled Down1" = Table.FillDown(#"Added Custom",{"Số tài khoản NH"}),
        #"Changed Type1" = Table.TransformColumnTypes(#"Filled Down1",{{"Số tài khoản NH", type text}}),
        #"Added Custom1" = Table.AddColumn(#"Changed Type1", "Số dư đầu file", each if [Ngày giao dịch] = "Số dư đầu kỳ" then [Số tham chiếu] else null),
        #"Filled Down2" = Table.FillDown(#"Added Custom1",{"Số dư đầu file"}),
        #"Removed Top Rows" = Table.Skip(#"Filled Down2",10),
        #"Removed Bottom Rows" = Table.RemoveLastN(#"Removed Top Rows",18),
        #"Added Custom2" = Table.AddColumn(#"Removed Bottom Rows", "Số dư đầu kỳ", each if Text.Start([Ngày đầu năm],5) = "01-01" and [Ngày giao dịch] ="Ngày giao dịch" then [Số dư đầu file] else 0),
        #"Changed Type2" = Table.TransformColumnTypes(#"Added Custom2",{{"Số dư đầu kỳ", type number}, {"Ngày hạch toán", type date}}),
        #"Changed Type with Locale" = Table.TransformColumnTypes(#"Changed Type2", {{"Ngày hạch toán", type date}}, "en-GB"),
        #"Replaced Value" = Table.ReplaceValue(#"Changed Type with Locale","Số tiền ghi có",null,Replacer.ReplaceValue,{"Phát sinh tăng tiền"}),
        #"Changed Type3" = Table.TransformColumnTypes(#"Replaced Value",{{"Phát sinh tăng tiền", type number}}),
        #"Replaced Value1" = Table.ReplaceValue(#"Changed Type3","Số tiền ghi nợ",null,Replacer.ReplaceValue,{"Phát sinh giảm tiền"}),
        #"Changed Type4" = Table.TransformColumnTypes(#"Replaced Value1",{{"Phát sinh giảm tiền", type number}}),
        #"Replaced Value2" = Table.ReplaceValue(#"Changed Type4","Ngày giao dịch",null,Replacer.ReplaceValue,{"Ngày giao dịch"}),
        #"Changed Type5" = Table.TransformColumnTypes(#"Replaced Value2",{{"Ngày giao dịch", type date}}),
        #"Changed Type with Locale1" = Table.TransformColumnTypes(#"Changed Type5", {{"Ngày giao dịch", type date}}, "en-GB"),
        #"Removed Columns" = Table.RemoveColumns(#"Changed Type with Locale1",{"Ngày đầu năm", "Số dư đầu file"}),
        #"Renamed Columns1" = Table.RenameColumns(#"Removed Columns",{{"Số tài khoản NH", "Số tài khoản NH 1"}}),
        #"Added Index" = Table.AddIndexColumn(#"Renamed Columns1", "Index", 1, 1, Int64.Type),
        #"Renamed Columns2" = Table.RenameColumns(#"Added Index",{{"Index", "Số chứng từ (STT bank)"}})
    in
        #"Renamed Columns2"
in
    Source