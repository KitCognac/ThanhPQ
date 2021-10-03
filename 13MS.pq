let
    Source = (Parameter13 as binary) => let
        Source = Excel.Workbook(Parameter13, null, true),
        #"Báo cáo_Sheet" = Source{[Item="Báo cáo",Kind="Sheet"]}[Data],
        #"Renamed Columns" = Table.RenameColumns(#"Báo cáo_Sheet",{{"Column1", "Chi nhánh"}, {"Column2", "Ngày hạch toán gốc"}, {"Column3", "Số chứng từ"}, {"Column4", "Mã mục thu/chi"}, {"Column5", "Tên mục thu/chi"}, {"Column6", "Mã AP"}, {"Column7", "Diễn giải"}, {"Column8", "Diễn giải riêng"}, {"Column9", "Tài khoản"}, {"Column10", "TK đối ứng"}, {"Column11", "Phát sinh tăng tiền"}, {"Column12", "Phát sinh giảm tiền"}, {"Column13", "Dư nợ"}, {"Column14", "Dư có"}, {"Column15", "Mã đối tượng"}, {"Column16", "Tên đối tượng"}, {"Column17", "Mã công trình"}, {"Column18", "Tên công trình"}, {"Column19", "Ngày bán vé"}, {"Column20", "Tuyến"}, {"Column21", "Tên tài khoản"}, {"Column22", "GNT ngân hàng"}, {"Column23", "Ghi chú"}, {"Column24", "Tên nhân viên"}, {"Column25", "CN được chi hộ"}, {"Column26", "Ngày hóa đơn"}, {"Column27", "Số hóa đơn"}, {"Column28", "Dư nợ ngoại tệ"}, {"Column29", "Dư có ngoại tệ"}, {"Column30", "Nghiệp vụ"}, {"Column31", "Loại chứng từ"}, {"Column32", "Mã nhân viên"}, {"Column33", "Mã khoản mục CP"}, {"Column34", "Tên khoản mục CP"}, {"Column35", "Mã đối tượng THCP"}, {"Column36", "Tên đối tượng THCP"}, {"Column37", "Hợp đồng bán"}, {"Column38", "Đơn hàng"}, {"Column39", "Hợp đồng mua"}, {"Column40", "Mã thống kê"}, {"Column41", "Tên thống kê"}, {"Column42", "Mã đơn vị"}, {"Column43", "Tên đơn vị"}, {"Column44", "CP hợp lý/không hợp lý"}}),
        #"Added Custom" = Table.AddColumn(#"Renamed Columns", "Kỳ xuất báo cáo", each if Text.Middle([Chi nhánh],25,4)="Năm " then "Từ ngày 01/01/"&Text.End([Chi nhánh],4)&" đến ngày 31/12/"&Text.End([Chi nhánh],4) else if
    Text.Start ([Chi nhánh],10 )="Loại tiền:"
    then Text.End ([Chi nhánh],38) else null),
        #"Filled Down" = Table.FillDown(#"Added Custom",{"Kỳ xuất báo cáo"}),
        #"Added Custom1" = Table.AddColumn(#"Filled Down", "Số dư đầu kỳ", each if [Diễn giải riêng] ="Số dư đầu kỳ" then [Dư nợ] else 0),
        #"Added Custom2" = Table.AddColumn(#"Added Custom1", "Số dư cuối kỳ", each if [Diễn giải riêng] ="Số dư cuối kỳ" then [Dư nợ] else 0),
        #"Removed Top Rows" = Table.Skip(#"Added Custom2",3),
        #"Added Custom3" = Table.AddColumn(#"Removed Top Rows", "Ngày hạch toán", each if [Diễn giải riêng] = "Số dư đầu kỳ" then Text.Middle ([Kỳ xuất báo cáo],8,10) else if [Diễn giải riêng] = "Số dư cuối kỳ" then Text.End ([Kỳ xuất báo cáo],Text.Length ([Kỳ xuất báo cáo])-28) else [Ngày hạch toán gốc]),
        #"Changed Type with Locale" = Table.TransformColumnTypes(#"Added Custom3", {{"Ngày hạch toán", type date}}, "en-GB"),
        #"Changed Type" = Table.TransformColumnTypes(#"Changed Type with Locale",{{"Mã mục thu/chi", type text}, {"Mã AP", type text}, {"Phát sinh tăng tiền", type number}, {"Phát sinh giảm tiền", type number}, {"Dư nợ", type number}, {"Dư có", type number}, {"Mã đối tượng", type text}, {"Mã công trình", type text}, {"Số hóa đơn", type text}, {"Dư nợ ngoại tệ", type number}, {"Dư có ngoại tệ", type number}, {"Mã đơn vị", type text}, {"Số dư đầu kỳ", type number}, {"Số dư cuối kỳ", type number}}),
        #"Removed Bottom Rows" = Table.RemoveLastN(#"Changed Type",1),
        #"Filtered Rows" = Table.SelectRows(#"Removed Bottom Rows", each ([Diễn giải riêng] <> "Cộng")),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Ngày hạch toán gốc"}),
        #"Renamed Columns2" = Table.RenameColumns(#"Removed Columns",{{"Tài khoản", "Tài khoản hạch toán"}}),
        #"Removed Columns1" = Table.RemoveColumns(#"Renamed Columns2",{"Dư có", "Mã công trình", "Tên công trình", "Ngày bán vé", "Tuyến", "Tên tài khoản", "Mã đối tượng THCP", "Tên đối tượng THCP", "Hợp đồng bán", "Đơn hàng"}),
        #"Renamed Columns3" = Table.RenameColumns(#"Removed Columns1",{{"Số chứng từ", "Số chứng từ (STT bank)"}, {"Chi nhánh", "Chi nhánh_Misa"},{"Tên đơn vị", "Tên đơn vị_Misa"}, {"Mã đơn vị", "Mã đơn vị_Misa"}, {"Mã đối tượng", "Mã đối tượng_Misa"}, {"Tên đối tượng", "Tên đối tượng_Misa"}}),
        #"Removed Columns2" = Table.RemoveColumns(#"Renamed Columns3",{"Column45"})
    in
        #"Removed Columns2"
in
    Source