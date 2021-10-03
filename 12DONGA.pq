let
    Source = (Parameter12 as binary) => let
        Source = Excel.Workbook(Parameter12, null, true),
        #"SAO KÊ_Sheet" = Source{0}[Data],
        #"Renamed Columns" = Table.RenameColumns(#"SAO KÊ_Sheet",{{"Column2", "STT"}, {"Column3", "Thời gian"}, {"Column4", "Địa điểm"}, {"Column5", "Chi tiết giao dịch"}, {"Column6", "Ghi chú giao dịch"}, {"Column7", "Tăng"}, {"Column8", "Giảm"}}),
        #"Added Custom" = Table.AddColumn(#"Renamed Columns", "Ngày đầu năm", each if Text.Middle ([STT] , 49,5) ="01-01" then "01-01-"&Text.Middle ([STT] , 55,4) else null),
        #"Filled Down" = Table.FillDown(#"Added Custom",{"Ngày đầu năm"}),
        #"Added Custom1" = Table.AddColumn(#"Filled Down", "Số tài khoản ngân hàng", each if Text.Start([STT],14) ="CÁC GIAO DỊCH " then Text.Middle ([STT],28,Text.Length([STT])-67) else null),
        #"Filled Down1" = Table.FillDown(#"Added Custom1",{"Số tài khoản ngân hàng"}),
        #"Added Custom2" = Table.AddColumn(#"Filled Down1", "Số dư đầu kỳ", each if Text.Start ([STT],13)="Số dư đầu kỳ:" and Text.Start([Ngày đầu năm],5)="01-01" then Text.End([STT],Text.Length([STT])-14) else 0),
        #"Added Custom3" = Table.AddColumn(#"Added Custom2", "Ngày hạch toán", each if Text.Start([STT],13)="Số dư đầu kỳ:"and Text.Start([Ngày đầu năm],5)="01-01" then [Ngày đầu năm] else [Thời gian]),
        #"Removed Top Rows" = Table.Skip(#"Added Custom3",4),
        #"Removed Bottom Rows" = Table.RemoveLastN(#"Removed Top Rows",1),
        #"Added Custom4" = Table.AddColumn(#"Removed Bottom Rows", "Phát sinh tăng tiền", each if [Tăng] = null then 0 else Text.End([Tăng],Text.Length([Tăng])-1)),
        #"Added Custom5" = Table.AddColumn(#"Added Custom4", "Phát sinh giảm tiền", each if [Giảm] = null or [Giảm] ="" then 0 else Text.End([Giảm],Text.Length([Giảm])-1)),
        #"Removed Columns" = Table.RemoveColumns(#"Added Custom5",{"Column1", "Column9", "Column10"}),
        #"Renamed Columns1" = Table.RenameColumns(#"Removed Columns",{{"Chi tiết giao dịch", "Diễn giải"}}),
        #"Removed Columns1" = Table.RemoveColumns(#"Renamed Columns1",{"Ngày đầu năm"}),
        #"Changed Type" = Table.TransformColumnTypes(#"Removed Columns1",{{"Số tài khoản ngân hàng", type text}, {"Số dư đầu kỳ", type number}}),
        #"Changed Type with Locale" = Table.TransformColumnTypes(#"Changed Type", {{"Ngày hạch toán", type date}}, "en-GB"),
        #"Changed Type1" = Table.TransformColumnTypes(#"Changed Type with Locale",{{"Phát sinh tăng tiền", type number}, {"Phát sinh giảm tiền", type number}}),
        #"Renamed Columns2" = Table.RenameColumns(#"Changed Type1",{{"Số tài khoản ngân hàng", "Số tài khoản NH"}}),
        #"Added Index" = Table.AddIndexColumn(#"Renamed Columns2", "Index", 1, 1, Int64.Type),
        #"Renamed Columns3" = Table.RenameColumns(#"Added Index",{{"Index", "Số chứng từ (STT bank)"}})
    in
        #"Renamed Columns3"
in
    Source