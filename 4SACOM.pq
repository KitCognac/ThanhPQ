let
    Source = (Parameter4 as binary) => let
        Source = Excel.Workbook(Parameter4, null, true),
        ViewStatementUX1 = Source{0}[Data],
        #"Renamed Columns" = Table.RenameColumns(ViewStatementUX1,{{"Column1", "STT"}, {"Column2", "Số giao dịch"}, {"Column4", "Ngày giao dịch"}, {"Column6", "Ngày hiệu lực"}, {"Column8", "Diễn giải"}, {"Column10", "Phát sinh giảm tiền"}, {"Column12", "Phát sinh tăng tiền"}, {"Column13", "Số dư hiện tại"}}),
        #"Added Custom" = Table.AddColumn(#"Renamed Columns", "Ngày đầu năm", each if Text.Start ([Column5],5) = "01-01" then [Column5] else null),
        #"Filled Down" = Table.FillDown(#"Added Custom",{"Ngày đầu năm"}),
        #"Changed Type" = Table.TransformColumnTypes(#"Filled Down",{{"Ngày đầu năm", type text}}),
        #"Added Custom1" = Table.AddColumn(#"Changed Type", "Ngày hạch toán", each if [STT] = "STT" then [Ngày đầu năm] else [Ngày hiệu lực]),
        #"Changed Type with Locale1" = Table.TransformColumnTypes(#"Added Custom1", {{"Ngày hạch toán", type date}}, "en-GB"),
        #"Added Custom2" = Table.AddColumn(#"Changed Type with Locale1", "Số đầu năm", each if [Column3] ="Số dư đầu kỳ:" then [Column5] else null),
        #"Filled Down1" = Table.FillDown(#"Added Custom2",{"Số đầu năm"}),
        #"Added Custom3" = Table.AddColumn(#"Filled Down1", "Số dư đầu kỳ", each if [STT] = "STT" and  Text.Start([Ngày đầu năm],5) ="01-01" then [Số đầu năm] else 0),
        #"Changed Type with Locale" = Table.TransformColumnTypes(#"Added Custom3", {{"Số dư đầu kỳ", type number}}, "vi-VN"),
        #"Added Custom4" = Table.AddColumn(#"Changed Type with Locale", "Số tài khoản NH", each if [Column3] = "Số Tài khoản:" then [Column5] else null),
        #"Filled Down2" = Table.FillDown(#"Added Custom4",{"Số tài khoản NH"}),
        #"Changed Type1" = Table.TransformColumnTypes(#"Filled Down2",{{"Số tài khoản NH", type text}}),
        #"Removed Top Rows" = Table.Skip(#"Changed Type1",11),
        #"Removed Bottom Rows" = Table.RemoveLastN(#"Removed Top Rows",2),
        #"Replaced Value" = Table.ReplaceValue(#"Removed Bottom Rows","Số tiền rút",null,Replacer.ReplaceValue,{"Phát sinh giảm tiền"}),
        #"Replaced Value1" = Table.ReplaceValue(#"Replaced Value","Số tiền gửi",null,Replacer.ReplaceValue,{"Phát sinh tăng tiền"}),
        #"Replaced Value2" = Table.ReplaceValue(#"Replaced Value1","Số dư hiện tại",null,Replacer.ReplaceValue,{"Số dư hiện tại"}),
        #"Removed Columns" = Table.RemoveColumns(#"Replaced Value2",{"Column3", "Column5", "Column7", "Column9", "Column11", "Column14", "Column15", "Column16", "Column17", "Ngày đầu năm", "Số đầu năm"}),
        #"Changed Type with Locale2" = Table.TransformColumnTypes(#"Removed Columns", {{"Phát sinh giảm tiền", type number}}, "vi-VN"),
        #"Changed Type with Locale3" = Table.TransformColumnTypes(#"Changed Type with Locale2", {{"Phát sinh tăng tiền", type number}}, "vi-VN"),
        #"Changed Type with Locale4" = Table.TransformColumnTypes(#"Changed Type with Locale3", {{"Số dư hiện tại", type number}}, "vi-VN"),
        #"Added Index" = Table.AddIndexColumn(#"Changed Type with Locale4", "Index", 1, 1, Int64.Type),
        #"Renamed Columns1" = Table.RenameColumns(#"Added Index",{{"Index", "Số chứng từ (STT bank)"}}),
        Custom1 = Table.Buffer(#"Renamed Columns1"),
        #"Added Custom5" = Table.AddColumn(Custom1, "Hình thức", each 
    if [Số dư đầu kỳ] >0 then "SDĐK" else
    if [STT] = "STT" then "Tên dòng" else
    if Text.Contains([Diễn giải],"RUT TM TK",Comparer.OrdinalIgnoreCase) then "Rút tiền mặt nhập quỹ" else
    if Text.Contains([Diễn giải],"TTT TT GIAO DICH ECOM NOI DIA",Comparer.OrdinalIgnoreCase) then "Napas" else 
    if Text.Contains([Diễn giải],"THU PHI CK GIAO DICH",Comparer.OrdinalIgnoreCase) then "Phí CKGD" else if Text.Contains([Diễn giải],"THU PHI XLGD",Comparer.OrdinalIgnoreCase) then "Phí XLGD" else if Text.Contains([Diễn giải],"PHI CK ",Comparer.OrdinalIgnoreCase) then "Phí CK" else if Text.Contains([Diễn giải],"PHI",Comparer.OrdinalIgnoreCase) and Text.Contains([Diễn giải],"UNC",Comparer.OrdinalIgnoreCase) then "Phí CK"else if Text.Contains([Diễn giải],"PHI",Comparer.OrdinalIgnoreCase) and Text.Contains([Diễn giải],",QLTK",Comparer.OrdinalIgnoreCase) then "Phí QLTK" else 
    if Text.Contains([Diễn giải],"LAI NHAP VON",Comparer.OrdinalIgnoreCase) then "Lãi tiền gửi thanh toán" else
    if Text.Contains([Diễn giải],"TONG CTY CP BAO HIEM BAO LONG BT XE",Comparer.OrdinalIgnoreCase) then "Tiền bồi thường xe từ TCT BH Bảo Long" else 
    if [Phát sinh giảm tiền] <> null then "CK tiền đi" else
    null),
        MA_DON_VI = Table.AddColumn(#"Added Custom5", "MA_DON_VI", each 
    let 
    _A = [Diễn giải], 
    _B = [Hình thức],
    _tbl = Table.SelectRows(DMVP, each Text.Contains(_A, [Mã đối tượng]) and (_B = "" or _B = null) ) 
    in try _tbl[Mã đối tượng]{0} otherwise null ),
        COUNT_MDV = Table.AddColumn(MA_DON_VI, "COUNT_MA_DON_VI", each 
    let 
    _A = [Diễn giải], 
    _B = [Hình thức],
    _tbl = Table.SelectRows(DMVP, each Text.Contains(_A, [Mã đối tượng]) and (_B = "" or _B = null) ) 
    in try Table.RowCount(_tbl) otherwise null )
    in
        COUNT_MDV
in
    Source