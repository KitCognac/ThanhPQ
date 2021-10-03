let
    Source = (Parameter2 as binary) => let
        Source = Excel.Workbook(Parameter2, null, true),
        RPT_LSGD1 = Source{0}[Data],
        #"Renamed Columns" = Table.RenameColumns(RPT_LSGD1,{{"Column2", "Ngày hiệu lực"}, {"Column4", "Phát sinh giảm tiền"}, {"Column5", "Phát sinh tăng tiền"}, {"Column6", "Số dư"}, {"Column7", "Diễn giải"}}),
        #"Added Custom" = Table.AddColumn(#"Renamed Columns", "Số tài khoản NH", each if Text.Start ([Column3] ,13) = "Số tài khoản:" then Text.End ([Column3] , Text.Length ([Column3])-14) else null),
        #"Filled Down" = Table.FillDown(#"Added Custom",{"Số tài khoản NH"}),
        #"Added Custom1" = Table.AddColumn(#"Filled Down", "Ngày đầu năm", each if Text.Start ([Column3] , 29 ) = "Ngày giao dịch Từ ngày: 01/01" then "01/01/"&Text.Middle ([Column3] , 30,4) else null),
        #"Filled Down1" = Table.FillDown(#"Added Custom1",{"Ngày đầu năm"}),
        #"Changed Type1" = Table.TransformColumnTypes(#"Filled Down1",{{"Ngày đầu năm", type text}}),
        #"Added Custom2" = Table.AddColumn(#"Changed Type1", "Ngày hạch toán", each if [Ngày hiệu lực] = "Dư đầu" then [Ngày đầu năm] else Text.Start ([Ngày hiệu lực] , 10)),
        #"Changed Type with Locale" = Table.TransformColumnTypes(#"Added Custom2", {{"Ngày hạch toán", type date}}, "en-GB"),
        #"Added Custom3" = Table.AddColumn(#"Changed Type with Locale", "Số dư đầu kỳ", each if [Ngày hiệu lực] = "Dư đầu" and Text.Start([Ngày đầu năm],5)="01/01" then [Số dư] else 0),
        #"Changed Type" = Table.TransformColumnTypes(#"Added Custom3",{{"Số dư đầu kỳ", type number}, {"Phát sinh giảm tiền", type number}, {"Phát sinh tăng tiền", type number}, {"Số dư", type number}, {"Số tài khoản NH", type text}}),
        #"Removed Top Rows" = Table.Skip(#"Changed Type",12),
        #"Removed Bottom Rows" = Table.RemoveLastN(#"Removed Top Rows",6),
        #"Removed Columns" = Table.RemoveColumns(#"Removed Bottom Rows",{"Column1", "Column3", "Column8", "Column9", "Ngày đầu năm"}),
        #"Added Index" = Table.AddIndexColumn(#"Removed Columns", "Index", 1, 1, Int64.Type),
        #"Renamed Columns1" = Table.RenameColumns(#"Added Index",{{"Index", "Số chứng từ (STT bank)"}})
    in
        #"Renamed Columns1"
in
    Source