let
    Source = (Parameter9 as binary) => let
        Source = Excel.Workbook(Parameter9, null, true),
        #"LICH SU GIAO DICH_Sheet" = Source{0}[Data],
        #"Renamed Columns" = Table.RenameColumns(#"LICH SU GIAO DICH_Sheet",{{"Column1", "STT/No."}, {"Column2", "Ngày giao dịch/ Transaction date"}, {"Column3", "Diễn giải"}, {"Column4", "Phát sinh giảm tiền"}, {"Column5", "Phát sinh tăng tiền"}, {"Column6", "Số dư TK/ Account Balance"}, {"Column7", "Số giao dịch/ Transaction number"}, {"Column8", "MTID/ CITAD"}}),
        #"Replaced Value3" = Table.ReplaceValue(#"Renamed Columns","Ngày giao dịch/ Transaction date",null,Replacer.ReplaceValue,{"Ngày giao dịch/ Transaction date"}),
        #"Duplicated Column" = Table.DuplicateColumn(#"Replaced Value3", "Ngày giao dịch/ Transaction date", "Ngày giao dịch/ Transaction date - Copy"),
        #"Changed Type" = Table.TransformColumnTypes(#"Duplicated Column",{{"Ngày giao dịch/ Transaction date - Copy", type text}}),
        #"Added Custom" = Table.AddColumn(#"Changed Type", "Số tài khoản NH", each if [#"Ngày giao dịch/ Transaction date - Copy"] = "Số tài khoản/ Account No.:" then [Diễn giải] else null),
        #"Filled Down" = Table.FillDown(#"Added Custom",{"Số tài khoản NH"}),
        #"Changed Type1" = Table.TransformColumnTypes(#"Filled Down",{{"Diễn giải", type text}}),
        #"Added Custom1" = Table.AddColumn(#"Changed Type1", "Ngày đầu năm", each if Text.Middle ([Diễn giải],5,5) ="01-01" then "01-01-"& Text.Start([Diễn giải],4) else null),
        #"Filled Down1" = Table.FillDown(#"Added Custom1",{"Ngày đầu năm"}),
        #"Added Custom2" = Table.AddColumn(#"Filled Down1", "Số đầu năm", each if [#"Ngày giao dịch/ Transaction date - Copy"] = "Số dư đầu kỳ/Opening Balance:" then [Diễn giải] else null),
        #"Filled Down2" = Table.FillDown(#"Added Custom2",{"Số đầu năm"}),
        #"Added Custom3" = Table.AddColumn(#"Filled Down2", "Ngày hạch toán", each if [#"STT/No."]= "STT/No." and Text.Start([Ngày đầu năm],5)= "01-01" then [Ngày đầu năm] else Text.Start([#"Ngày giao dịch/ Transaction date - Copy"],10)),
        #"Added Custom4" = Table.AddColumn(#"Added Custom3", "Số dư đầu kỳ", each if [#"STT/No."]="STT/No." and Text.Start([Ngày đầu năm],5)= "01-01" then [Số đầu năm] else 0),
        #"Removed Top Rows" = Table.Skip(#"Added Custom4",24),
        #"Replaced Value" = Table.ReplaceValue(#"Removed Top Rows","Nợ/ Debit",null,Replacer.ReplaceValue,{"Phát sinh giảm tiền"}),
        #"Replaced Value1" = Table.ReplaceValue(#"Replaced Value","Có / Credit",null,Replacer.ReplaceValue,{"Phát sinh tăng tiền"}),
        #"Replaced Value2" = Table.ReplaceValue(#"Replaced Value1","Số dư TK/ Account Balance",null,Replacer.ReplaceValue,{"Số dư TK/ Account Balance"}),
        #"Changed Type2" = Table.TransformColumnTypes(#"Replaced Value2",{{"Phát sinh giảm tiền", type number}, {"Phát sinh tăng tiền", type number}, {"Số dư TK/ Account Balance", type number}, {"Số tài khoản NH", type text}, {"Ngày đầu năm", type date}, {"Số đầu năm", type number}}),
        #"Changed Type with Locale" = Table.TransformColumnTypes(#"Changed Type2", {{"Ngày hạch toán", type date}}, "en-GB"),
        #"Changed Type3" = Table.TransformColumnTypes(#"Changed Type with Locale",{{"Số dư đầu kỳ", type number}}),
        #"Removed Columns" = Table.RemoveColumns(#"Changed Type3",{"Ngày giao dịch/ Transaction date - Copy", "Ngày đầu năm", "Số đầu năm"}),
        #"Added Index" = Table.AddIndexColumn(#"Removed Columns", "Index", 1, 1, Int64.Type),
        #"Renamed Columns1" = Table.RenameColumns(#"Added Index",{{"Index", "Số chứng từ (STT bank)"}}),
     #"Added Custom5" = Table.AddColumn(#"Renamed Columns1", "Hình thức", each
    if [Số dư đầu kỳ] > 0 then "SDĐK" else
    if [#"STT/No."]="STT/No." then "Tên dòng" else 
    if Text.Contains([Diễn giải],"Tra lai tai khoan DDA",Comparer.OrdinalIgnoreCase) then "Lãi tiền gửi thanh toán" else
    if Text.Contains([Diễn giải],"Tra no TK vay",Comparer.OrdinalIgnoreCase) then "Trả nợ TK vay" else 
    if Text.Contains([Diễn giải],"THU LAI",Comparer.OrdinalIgnoreCase) then "Trả nợ TK vay" else
    if Text.Contains([Diễn giải],"THU MOT PHAN NO GOC",Comparer.OrdinalIgnoreCase) then "Trả nợ TK vay" else
    if Text.Contains([Diễn giải],"THU NO GOC",Comparer.OrdinalIgnoreCase) then "Trả nợ TK vay" else
    if Text.Contains([Diễn giải],"THU GOC ",Comparer.OrdinalIgnoreCase) then "Trả nợ TK vay" else
    if Text.Contains([Diễn giải],"LAI DEN HAN",Comparer.OrdinalIgnoreCase) then "Trả nợ TK vay" else  
    if Text.Contains([Diễn giải],"Purchase:ECOM",Comparer.OrdinalIgnoreCase) and [Số tài khoản NH] ="117000140190" then "Cyber" else
    if Text.Contains([Diễn giải],"Phi GD:",Comparer.OrdinalIgnoreCase) and [Số tài khoản NH] ="117000140190" then "Phí Cyber" else
    if Text.Contains([Diễn giải],"TTTT",Comparer.OrdinalIgnoreCase) and Text.Contains([Diễn giải],"Bao co",Comparer.OrdinalIgnoreCase) and [Số tài khoản NH] ="117000140190" then "Ngân lượng" else
    if Text.Contains([Diễn giải],"TTTT",Comparer.OrdinalIgnoreCase) and Text.Contains([Diễn giải],"Thu phi",Comparer.OrdinalIgnoreCase) and [Số tài khoản NH] ="117000140190" then "Phí Ngân lượng" else
    if Text.Contains([Diễn giải],"VNPAY",Comparer.OrdinalIgnoreCase) and [Số tài khoản NH] ="117000140190" then "VNPAY" else
    if Text.Contains([Diễn giải],"MoMo thanh toan futabuslines",Comparer.OrdinalIgnoreCase) and [Số tài khoản NH] ="117000140190" then "MOMO" else
    if Text.Contains([Diễn giải],"m4becomfutabuslines",Comparer.OrdinalIgnoreCase) and Text.Contains([Diễn giải],"Momo",Comparer.OrdinalIgnoreCase) and [Số tài khoản NH] ="117000140190" then "MOMO" else
    if Text.Contains([Diễn giải],"MoMo thanh toan futabus",Comparer.OrdinalIgnoreCase) and [Số tài khoản NH] ="117000140190"then "APPMOMO" else
    if Text.Contains([Diễn giải],"m4becomfutabus",Comparer.OrdinalIgnoreCase) and Text.Contains([Diễn giải],"Momo",Comparer.OrdinalIgnoreCase) and [Số tài khoản NH] ="117000140190" then "APPMOMO" else 
    if Text.Contains([Diễn giải],"CKRmNo:",Comparer.OrdinalIgnoreCase) and Text.Contains([Diễn giải],"Momo",Comparer.OrdinalIgnoreCase) and [Số tài khoản NH] ="117000140190" then "MOMO hoặc APP MOMO" else 
    if (Text.Contains([Diễn giải],"VIFUTA",Comparer.OrdinalIgnoreCase) or Text.Contains([Diễn giải],"VIFUATA",Comparer.OrdinalIgnoreCase)) and [Số tài khoản NH] ="117000140190" then "Ví Futa" else
    if Text.Contains([Diễn giải],"PMR",Comparer.OrdinalIgnoreCase) and [Số tài khoản NH] ="117000140190" then "ZALOPAY" else
    if Text.Contains([Diễn giải],"CT DEN",Comparer.OrdinalIgnoreCase) and [Số tài khoản NH] ="117000140190" then "CK Vietinbank" else
    if Text.Contains([Diễn giải],"THU HOI GD HOAN TRA",Comparer.OrdinalIgnoreCase) and [Số tài khoản NH] ="117000140190" then "Hoàn trả tiền khách mua vé" else  
    if Text.Contains([Diễn giải],"So GD: 902A2",Comparer.OrdinalIgnoreCase) then "Chuyển tiền sang TK khác hoặc phí chuyển tiền" else  
    if Text.Contains([Diễn giải],"CT DI:",Comparer.OrdinalIgnoreCase) then "Chuyển tiền sang TK khác hoặc phí chuyển tiền" else
    if [Số tài khoản NH] ="117000140190" and [Phát sinh tăng tiền]>0 then "KH mua vé CK Vietin hoặc khoản nhận tiền khác" else 
    if [Số tài khoản NH] ="115000130284" and Text.Contains([Diễn giải],"Thu phi duy tri dich vu IB",Comparer.OrdinalIgnoreCase) then "Phí duy trì TK IB" else
    if [Số tài khoản NH] ="115000130284" and [Diễn giải] ="VAT" then "VAT - Phí duy trì TK IB" else
    "Các khoản chưa phân loại")
    in
        #"Added Custom5"
in
    Source