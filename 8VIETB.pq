let
    Source = (Parameter8 as binary) => let
        Source = Excel.Workbook(Parameter8, null, true),
        VietBank_01012021_31012021_Sheet = Source{0}[Data],
        #"Renamed Columns" = Table.RenameColumns(VietBank_01012021_31012021_Sheet,{{"Column1", "Ngày"}, {"Column2", "Số giao dịch"}, {"Column3", "Ghi nợ"}, {"Column4", "Ghi có"}, {"Column5", "Số dư"}}),
        #"Duplicated Column" = Table.DuplicateColumn(#"Renamed Columns", "Ngày", "Ngày - Copy"),
        #"Changed Type" = Table.TransformColumnTypes(#"Duplicated Column",{{"Ngày - Copy", type text}}),
        #"Removed Bottom Rows" = Table.RemoveLastN(#"Changed Type",1),
        #"Added Custom4" = Table.AddColumn(#"Removed Bottom Rows", "Số tài khoản NH", each if [#"Ngày - Copy"] = "Số tài khoản:" then [Số giao dịch] else null),
        #"Filled Down2" = Table.FillDown(#"Added Custom4",{"Số tài khoản NH"}),
        #"Changed Type1" = Table.TransformColumnTypes(#"Filled Down2",{{"Số tài khoản NH", type text}}),
        #"Replaced Value" = Table.ReplaceValue(#"Changed Type1","'","",Replacer.ReplaceText,{"Số tài khoản NH"}),
        #"Removed Top Rows" = Table.Skip(#"Replaced Value",12),
        #"Added Custom3" = Table.AddColumn(#"Removed Top Rows", "Ngày hạch toán", each [Ngày]),
        #"Changed Type with Locale" = Table.TransformColumnTypes(#"Added Custom3", {{"Ngày hạch toán", type date}}, "en-GB"),
        #"Changed Type3" = Table.TransformColumnTypes(#"Changed Type with Locale",{{"Số dư", type number}, {"Ghi có", type number}, {"Ghi nợ", type number}}),
        #"Added Custom5" = Table.AddColumn(#"Changed Type3", "Phát sinh", each if [Ghi nợ] =null and [Ghi có] = null then null else if [Ghi nợ]=null then [Ghi có] else -1*[Ghi nợ]),
        #"Changed Type4" = Table.TransformColumnTypes(#"Added Custom5",{{"Phát sinh", type number}}),
        #"Filled Down3" = Table.FillDown(#"Changed Type4",{"Phát sinh"}),
        #"Added Custom6" = Table.AddColumn(#"Filled Down3", "Phát sinh tăng tiền", each if [Phát sinh] = null or [Phát sinh] >0 then [Phát sinh] else 0),
        #"Added Custom7" = Table.AddColumn(#"Added Custom6", "Phát sinh giảm tiền", each if [Phát sinh] > 0 or [Phát sinh] = null then 0 else -1*[Phát sinh]),
        #"Changed Type5" = Table.TransformColumnTypes(#"Added Custom7",{{"Phát sinh tăng tiền", type number}, {"Phát sinh giảm tiền", type number}}),
        #"Added Custom8" = Table.AddColumn(#"Changed Type5", "Diễn giải", each if Text.Start ([Số giao dịch],1) = "'" then null else [Số giao dịch]),
        #"Filled Down4" = Table.FillDown(#"Added Custom8",{"Ngày", "Số dư", "Ngày - Copy", "Ngày hạch toán"}),
        #"Filtered Rows" = Table.SelectRows(#"Filled Down4", each ([Diễn giải] <> null)),
        #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Ghi nợ", "Ghi có", "Ngày - Copy", "Phát sinh", "Số giao dịch"}),
        #"Renamed Columns1" = Table.RenameColumns(#"Removed Columns",{{"Số dư", "Số dư giao dịch"}}),
        #"Added Index" = Table.AddIndexColumn(#"Renamed Columns1", "Index", 1, 1, Int64.Type),
        #"Renamed Columns2" = Table.RenameColumns(#"Added Index",{{"Index", "Số chứng từ (STT bank)"}})
    in
        #"Renamed Columns2"
in
    Source