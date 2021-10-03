let
    Source = (Parameter6 as binary) => let
        Source = Excel.Workbook(Parameter6, null, true),
        Sheet2 = Source{0}[Data],
        #"Đổi tên tiêu đề" = Table.RenameColumns(Sheet2,{{"Column2", "STT"}, {"Column3", "SỐ THAM CHIẾU"}, {"Column4", "NGÀY GIAO DỊCH"}, {"Column5", "NGÀY HẠCH TOÁN 2"}, {"Column6", "Phát sinh giảm tiền"}, {"Column7", "Phát sinh tăng tiền"}, {"Column8", "Diễn giải"}}),
        #"Bỏ 3 dòng cuối file" = Table.RemoveLastN(#"Đổi tên tiêu đề",3),
        #"Thêm cột - Ngày đầu năm" = Table.AddColumn(#"Bỏ 3 dòng cuối file", "Ngày đầu năm", each if Text.Start([NGÀY HẠCH TOÁN 2] ,18)= "(From date): 01/01" then "01-01-"&Text.Trim(Text.End([NGÀY HẠCH TOÁN 2] ,5)) else null),
        #"Repeat lại ngày cột ngày đầu năm" = Table.FillDown(#"Thêm cột - Ngày đầu năm",{"Ngày đầu năm"}),
        #"Thêm cột Ngày hạch toán" = Table.AddColumn(#"Repeat lại ngày cột ngày đầu năm", "Ngày hạch toán", each if [STT] = "STT" then [Ngày đầu năm] else [NGÀY HẠCH TOÁN 2]),
        #"Thêm cột Số tài khoản NH" = Table.AddColumn(#"Thêm cột Ngày hạch toán", "Số tài khoản NH", each if [SỐ THAM CHIẾU] ="Số tài khoản " then Text.End ( [NGÀY GIAO DỊCH], Text.Length([NGÀY GIAO DỊCH])-18) else null),
        #"Repeat lại số tài khoản trên cột Số TK NH" = Table.FillDown(#"Thêm cột Số tài khoản NH",{"Số tài khoản NH"}),
        #"Định dạng cột số tài khoản NH" = Table.TransformColumnTypes(#"Repeat lại số tài khoản trên cột Số TK NH",{{"Số tài khoản NH", type text}}),
        #"Thêm cột Số đầu năm" = Table.AddColumn(#"Định dạng cột số tài khoản NH", "Số đầu năm", each if [SỐ THAM CHIẾU] = "Số dư đầu kỳ " then Text.End ([NGÀY GIAO DỊCH], Text.Length ([NGÀY GIAO DỊCH])-19) else null),
        #"Repeat số đầu năm" = Table.FillDown(#"Thêm cột Số đầu năm",{"Số đầu năm"}),
        #"Định dạng số cột số đầu năm" = Table.TransformColumnTypes(#"Repeat số đầu năm",{{"Số đầu năm", type number}}),
        #"Thêm cột số dư đầu kỳ" = Table.AddColumn(#"Định dạng số cột số đầu năm", "Số dư đầu kỳ", each if  [STT] ="STT" and Text.Start([Ngày đầu năm],5) = "01-01" then [Số đầu năm] else 0),
        #"Bỏ 22 dòng đầu tiên mỗi file" = Table.Skip(#"Thêm cột số dư đầu kỳ",22),
        #"Bỏ cột không cần" = Table.RemoveColumns(#"Bỏ 22 dòng đầu tiên mỗi file",{"Column1", "Số đầu năm", "Ngày đầu năm"}),
        #"Replaced Value" = Table.ReplaceValue(#"Bỏ cột không cần","NGÀY HẠCH TOÁN",null,Replacer.ReplaceValue,{"NGÀY HẠCH TOÁN 2"}),
        #"Replaced Value1" = Table.ReplaceValue(#"Replaced Value","GHI NỢ",null,Replacer.ReplaceValue,{"Phát sinh giảm tiền"}),
        #"Replaced Value2" = Table.ReplaceValue(#"Replaced Value1","GHI CÓ",null,Replacer.ReplaceValue,{"Phát sinh tăng tiền"}),
        #"Định dạng dữ liệu" = Table.TransformColumnTypes(#"Replaced Value2",{{"Số dư đầu kỳ", type number}, {"Phát sinh giảm tiền", type number}, {"Phát sinh tăng tiền", type number}}),
        #"Định dạng ngày trên cột Ngày hạch toán" = Table.TransformColumnTypes(#"Định dạng dữ liệu", {{"Ngày hạch toán", type date}}, "en-GB"),
        #"Removed Columns" = Table.RemoveColumns(#"Định dạng ngày trên cột Ngày hạch toán",{"NGÀY HẠCH TOÁN 2"}),
        #"Added Index" = Table.AddIndexColumn(#"Removed Columns", "Index", 1, 1, Int64.Type),
        #"Renamed Columns" = Table.RenameColumns(#"Added Index",{{"Index", "Số chứng từ (STT bank)"}})
    in
        #"Renamed Columns"
in
    Source