let
    Source = (Parameter10 as binary) => let
        Source = Excel.Workbook(Parameter10, null, true),
        #"SAO KÊ_Sheet" = Source{0}[Data],
        #"Renamed Columns" = Table.RenameColumns(#"SAO KÊ_Sheet",{{"Column1", "STT"}, {"Column2", "Số bút toán"}, {"Column3", "Ngày giá trị"}, {"Column4", "Phát sinh tăng tiền"}, {"Column5", "Phát sinh giảm tiền"}, {"Column6", "Diễn giải"}, {"Column7", "Số dư sau giao dịch"}, {"Column8", "Thời gian giao dịch"}}),
        #"Added Custom" = Table.AddColumn(#"Renamed Columns", "Số tài khoản NH", each if Text.Start ([STT] , 13)="Số tài khoản:" then Text.Middle ( [STT],14,8) else null),
        #"Filled Down" = Table.FillDown(#"Added Custom",{"Số tài khoản NH"}),
        #"Changed Type" = Table.TransformColumnTypes(#"Filled Down",{{"Số tài khoản NH", type text}}),
        #"Added Custom1" = Table.AddColumn(#"Changed Type", "Ngày đầu năm", each if Text.Start ([STT],16)= "Kỳ sao kê: 01/01" then "01-01-"&Text.Middle([STT],17,4) else null),
        #"Filled Down1" = Table.FillDown(#"Added Custom1",{"Ngày đầu năm"}),
        #"Added Custom2" = Table.AddColumn(#"Filled Down1", "Số đầu năm", each if Text.Start ([STT] ,13)="Số dư đầu kỳ:" then Text.Middle([STT],14,Text.Length([STT])-18) else null),
        #"Filled Down2" = Table.FillDown(#"Added Custom2",{"Số đầu năm"}),
        #"Added Custom3" = Table.AddColumn(#"Filled Down2", "Ngày hạch toán", each if [STT]="#" and Text.Start([Ngày đầu năm],5)="01-01" then [Ngày đầu năm] else [Ngày giá trị]),
        #"Added Custom4" = Table.AddColumn(#"Added Custom3", "Số dư đầu kỳ", each if [STT]="#"and Text.Start([Ngày đầu năm],5)="01-01" then [Số đầu năm] else 0),
        #"Removed Top Rows" = Table.Skip(#"Added Custom4",16),
        #"Replaced Value" = Table.ReplaceValue(#"Removed Top Rows","Credit Amount",null,Replacer.ReplaceValue,{"Phát sinh tăng tiền"}),
        #"Replaced Value1" = Table.ReplaceValue(#"Replaced Value","Debit Amount",null,Replacer.ReplaceValue,{"Phát sinh giảm tiền"}),
        #"Replaced Value2" = Table.ReplaceValue(#"Replaced Value1","Account.Balance",null,Replacer.ReplaceValue,{"Số dư sau giao dịch"}),
        #"Replaced Value3" = Table.ReplaceValue(#"Replaced Value2","Value Date",null,Replacer.ReplaceValue,{"Ngày hạch toán"}),
        #"Removed Columns" = Table.RemoveColumns(#"Replaced Value3",{"Column9", "Ngày đầu năm", "Số đầu năm"}),
        #"Changed Type1" = Table.TransformColumnTypes(#"Removed Columns",{{"Số dư đầu kỳ", type number}}),
        #"Changed Type with Locale" = Table.TransformColumnTypes(#"Changed Type1", {{"Ngày hạch toán", type date}}, "en-GB"),
        #"Changed Type2" = Table.TransformColumnTypes(#"Changed Type with Locale",{{"Số dư sau giao dịch", type number}, {"Phát sinh giảm tiền", type number}, {"Phát sinh tăng tiền", type number}}),
        #"Filtered Rows" = Table.SelectRows(#"Changed Type2", each ([Ngày giá trị] <> null)),
        #"Added Index" = Table.AddIndexColumn(#"Filtered Rows", "Index", 1, 1, Int64.Type),
        #"Renamed Columns1" = Table.RenameColumns(#"Added Index",{{"Index", "Số chứng từ (STT bank)"}})
    in
        #"Renamed Columns1"
in
    Source