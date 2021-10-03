let
    Source = (Parameter3 as binary) => let
        Source = Excel.Workbook(Parameter3, null, true),
        Sheet1_Sheet = Source{0}[Data],
        #"Changed Type" = Table.TransformColumnTypes(Sheet1_Sheet,{{"Column1", type text}}),
        #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{{"Column1", "STT"}, {"Column2", "Ngày giao dịch"}, {"Column3", "Số bút toán"}, {"Column4", "Phát sinh giảm tiền"}, {"Column5", "Phát sinh tăng tiền"}, {"Column6", "Số dư"}, {"Column7", "Diễn giải"}}),
        #"Added Custom4" = Table.AddColumn(#"Renamed Columns", "Ngày đầu năm", each if Text.Start(Text.End([Số bút toán],14),7)="01 - 01" then "01-01-"&Text.End([Số bút toán],4) else null),
        #"Filled Down2" = Table.FillDown(#"Added Custom4",{"Ngày đầu năm"}),
        #"Added Custom5" = Table.AddColumn(#"Filled Down2", "Số đầu năm", each if Text.Start([Ngày giao dịch],12) = "Số dư đầu kỳ" then [Ngày giao dịch] else null),
        #"Replaced Value" = Table.ReplaceValue(#"Added Custom5",".","",Replacer.ReplaceText,{"Số đầu năm"}),
        #"Replaced Value1" = Table.ReplaceValue(#"Replaced Value",",00","",Replacer.ReplaceText,{"Số đầu năm"}),
        #"Replaced Value2" = Table.ReplaceValue(#"Replaced Value1","Số dư đầu kỳ: ","",Replacer.ReplaceText,{"Số đầu năm"}),
        #"Changed Type2" = Table.TransformColumnTypes(#"Replaced Value2",{{"Số đầu năm", type number}}),
        #"Filled Down3" = Table.FillDown(#"Changed Type2",{"Số đầu năm"}),
        #"Added Custom" = Table.AddColumn(#"Filled Down3", "Ngày hạch toán", each if Text.Start(Text.From([Phát sinh tăng tiền])
    ,8)="Tháng 01" then [Ngày đầu năm] else [Ngày giao dịch]),
        #"Added Custom2" = Table.AddColumn(#"Added Custom", "Số dư đầu kỳ", each if Text.Start(Text.From([Phát sinh tăng tiền])
    ,8)="Tháng 01" then [Số đầu năm] else 0),
        #"Added Custom3" = Table.AddColumn(#"Added Custom2", "Số tài khoản NH", each if Text.Start(Text.From([Phát sinh tăng tiền]),13) = "Số tài khoản:" then Text.End(Text.From([Phát sinh tăng tiền]),Text.Length([Phát sinh tăng tiền])-14) else null),
        #"Filled Down1" = Table.FillDown(#"Added Custom3",{"Số tài khoản NH"}),
        #"Changed Type1" = Table.TransformColumnTypes(#"Filled Down1",{{"Số tài khoản NH", type text}}),
        #"Removed Top Rows" = Table.Skip(#"Changed Type1",7),
        #"Removed Bottom Rows" = Table.RemoveLastN(#"Removed Top Rows",1),
        // Thay thế giá trị những dòng có chữ tháng bằng null
        // 
        #"Thay thế giá trị tháng ở cột PS tăng tiền" = Table.ReplaceValue(#"Removed Bottom Rows",each [Phát sinh tăng tiền],each if Text.Start([Phát sinh tăng tiền],5)="Tháng" then null else [Phát sinh tăng tiền],Replacer.ReplaceValue,{"Phát sinh tăng tiền"}),
        #"Changed Type with Locale" = Table.TransformColumnTypes(#"Thay thế giá trị tháng ở cột PS tăng tiền", {{"Phát sinh giảm tiền", type number}}, "en-GB"),
        #"Changed Type with Locale1" = Table.TransformColumnTypes(#"Changed Type with Locale", {{"Phát sinh tăng tiền", type number}}, "en-GB"),
        #"Changed Type with Locale2" = Table.TransformColumnTypes(#"Changed Type with Locale1", {{"Số dư", type number}}, "en-GB"),
        #"Changed Type with Locale3" = Table.TransformColumnTypes(#"Changed Type with Locale2", {{"Ngày hạch toán", type date}}, "en-GB"),
        #"Changed Type with Locale5" = Table.TransformColumnTypes(#"Changed Type with Locale3", {{"Số dư đầu kỳ", type number}}, "vi-VN"),
        #"Added Index" = Table.AddIndexColumn(#"Changed Type with Locale5", "Index", 1, 1, Int64.Type),
        #"Renamed Columns1" = Table.RenameColumns(#"Added Index",{{"Index", "Số chứng từ (STT bank)"}}),
        #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns1",{"Số đầu năm", "Ngày đầu năm"})
    in
        #"Removed Columns"
in
    Source