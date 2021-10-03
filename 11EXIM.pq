let
    Source = (Parameter11 as binary) => let
        Source = Excel.Workbook(Parameter11, null, true),
        Test_Sheet = Source{0}[Data],
        #"Renamed Columns" = Table.RenameColumns(Test_Sheet,{{"Column2", "Ngày giờ giao dịch"}, {"Column3", "Ngày hiệu lực"}, {"Column4", "Phát sinh giảm tiền"}, {"Column5", "Phát sinh tăng tiền"}, {"Column6", "Số dư sau giao dịch"}, {"Column7", "Diễn giải"}, {"Column1", "STT"}}),
        #"Added Custom" = Table.AddColumn(#"Renamed Columns", "Số tài khoản NH", each if [STT] = "Số tài khoản / Account No.:" then [Phát sinh giảm tiền] else null),
        #"Filled Down" = Table.FillDown(#"Added Custom",{"Số tài khoản NH"}),
        #"Added Custom1" = Table.AddColumn(#"Filled Down", "Ngày đầu năm", each if [Phát sinh giảm tiền]= "01/04/2021" or Text.Start ([Phát sinh giảm tiền],5) = "01/01"  then "01-01-"&Text.End ([Phát sinh giảm tiền],4)  else null),
        #"Filled Down1" = Table.FillDown(#"Added Custom1",{"Ngày đầu năm"}),
        #"Added Custom2" = Table.AddColumn(#"Filled Down1", "Số dư đầu kỳ", each if [STT]="Số dư đầu kỳ / Previous Balance " and Text.Start([Ngày đầu năm],5) = "01-01" then [Số dư sau giao dịch] else 0),
        #"Added Custom3" = Table.AddColumn(#"Added Custom2", "Ngày hạch toán", each if [STT]="Số dư đầu kỳ / Previous Balance " and Text.Start([Ngày đầu năm],5) = "01-01" then [Ngày đầu năm] else Text.Start ([Ngày hiệu lực],10)),
        #"Removed Top Rows" = Table.Skip(#"Added Custom3",6),
        #"Removed Columns" = Table.RemoveColumns(#"Removed Top Rows",{"Column8", "Ngày đầu năm"}),
        #"Changed Type with Locale" = Table.TransformColumnTypes(#"Removed Columns", {{"Ngày hạch toán", type date}}, "en-GB"),
        #"Changed Type" = Table.TransformColumnTypes(#"Changed Type with Locale",{{"Số dư đầu kỳ", type number}, {"Số tài khoản NH", type text}, {"Số dư sau giao dịch", type number}, {"Phát sinh tăng tiền", type number}, {"Phát sinh giảm tiền", type number}}),
        #"Removed Bottom Rows" = Table.RemoveLastN(#"Changed Type",5),
        #"Added Index" = Table.AddIndexColumn(#"Removed Bottom Rows", "Index", 1, 1, Int64.Type),
        #"Renamed Columns1" = Table.RenameColumns(#"Added Index",{{"Index", "Số chứng từ (STT bank)"}})
    in
        #"Renamed Columns1"
in
    Source