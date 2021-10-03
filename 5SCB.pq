let
    Source = (Parameter5 as binary) => let
        Source = Excel.Workbook(Parameter5, null, true),
        #"SCB-112101.SCB.01-T01_Sheet" = Source{0}[Data],
        #"Đổi tên các cột tiêu đề" = Table.RenameColumns(#"SCB-112101.SCB.01-T01_Sheet",{{"Column1", "Ngày hạch toán"}, {"Column2", "Diễn giải"}, {"Column3", "SO GIAO DICH"}, {"Column4", "LOAI TIEN"}, {"Column5", "Phát sinh giảm tiền"}, {"Column6", "Phát sinh tăng tiền"}, {"Column7", "SO DU"}}),
        #"Thêm cột - Số TK NH" = Table.AddColumn(#"Đổi tên các cột tiêu đề", "Số tài khoản NH", each if [Ngày hạch toán] = 1370105338560000 then 1370105338560001 else if [Ngày hạch toán] = 2270112363520000 then 2270112363520001 else if [Ngày hạch toán] = 430107036200000 then 430107036200001 else if [Diễn giải] = "CTY CP XE KHACH PHUONG TRANG FUTABUSLINES" then [Ngày hạch toán] else null),
        #"Filled Down" = Table.FillDown(#"Thêm cột - Số TK NH",{"Số tài khoản NH"}),
        #"Changed Type" = Table.TransformColumnTypes(#"Filled Down",{{"Số tài khoản NH", type text}}),
        #"Thêm cột - số đầu năm" = Table.AddColumn(#"Changed Type", "Số đầu năm", each if [Ngày hạch toán] = "SO DU DAU KY" then [SO GIAO DICH] else null),
        #"Filled Down1" = Table.FillDown(#"Thêm cột - số đầu năm",{"Số đầu năm"}),
        #"Removed Top Rows" = Table.Skip(#"Filled Down1",3),
        #"Replaced Value" = Table.ReplaceValue(#"Removed Top Rows","NGAY",null,Replacer.ReplaceValue,{"Ngày hạch toán"}),
        #"Replaced Value1" = Table.ReplaceValue(#"Replaced Value","TIEN RA",null,Replacer.ReplaceValue,{"Phát sinh giảm tiền"}),
        #"Replaced Value2" = Table.ReplaceValue(#"Replaced Value1","TIEN VAO",null,Replacer.ReplaceValue,{"Phát sinh tăng tiền"}),
        #"Replaced Value3" = Table.ReplaceValue(#"Replaced Value2","SO DU",null,Replacer.ReplaceValue,{"SO DU"}),
        #"Changed Type with Locale" = Table.TransformColumnTypes(#"Replaced Value3", {{"Ngày hạch toán", type date}}, "en-GB"),
        #"Filtered Rows" = Table.SelectRows(#"Changed Type with Locale", each ([Ngày hạch toán] <> null)),
        #"Renamed Columns" = Table.RenameColumns(#"Filtered Rows",{{"Số đầu năm", "Số dư đầu file"}, {"Số tài khoản NH", "Số tài khoản NH 1"}}),
        #"Added Index" = Table.AddIndexColumn(#"Renamed Columns", "Index", 1, 1, Int64.Type),
        #"Renamed Columns1" = Table.RenameColumns(#"Added Index",{{"Index", "Số chứng từ (STT bank)"}})
    in
        #"Renamed Columns1"
in
    Source