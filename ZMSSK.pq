let
    Source = LibPQ("Table.XLFilesLoad")("FILE_LIST","CURRENT_FOLDER_PATH",true),
     #"Changed Type with Locale" = Table.TransformColumnTypes(Source, {{"Ngày hạch toán", type date}}, "en-GB"),
    #"Changed Type" = Table.TransformColumnTypes(#"Changed Type with Locale",{{"Số tài khoản NH", type text}, {"Số dư đầu kỳ", type number}, {"Phát sinh tăng tiền", type number}, {"Phát sinh giảm tiền", type number}, {"Số dư sau giao dịch", type number}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{{"FILE_ID", "Tên file tổng ngân hàng"}}),
    #"Trimmed Text" = Table.TransformColumns(#"Renamed Columns",{{"Số tài khoản NH", Text.Trim, type text}}),
    #"Removed Columns" = Table.RemoveColumns(#"Trimmed Text",{"Số dư sau giao dịch"}),
    #"Filtered Rows1" = Table.SelectRows(#"Removed Columns", each ([Ngày hạch toán] <> null)),
    #"Removed Columns1" = Table.RemoveColumns(#"Filtered Rows1",{"STT", "Ngay Giao dich", "Tu tai khoan", "Den tai khoan", "So Du truoc", "+/-", "So Tien Giao dich", "So Du", "Chi nhanh", "Trang thai", "Ngày giao dịch", "VERSION", "Số bút toán", "Ngày giá trị", "Thời gian giao dịch", "Thời gian", "Địa điểm", "Ghi chú giao dịch", "Tăng", "Giảm", "Số giao dịch", "SO GIAO DICH", "LOAI TIEN", "SO DU", "Column8", "Số tài khoản NH 1", "Số dư đầu file", "Số dư đầu kỳ tại ngày", "SỐ THAM CHIẾU", "NGÀY GIAO DỊCH", "Số tham chiếu", "Column6", "Ngày", "STT/No.", "Ngày giao dịch/ Transaction date", "Số dư TK/ Account Balance", "Số giao dịch/ Transaction number", "MTID/ CITAD", "Loai Tien", "Ngày hiệu lực"}),
    #"Renamed Columns1" = Table.RenameColumns(#"Removed Columns1",{{"DESC", "Loại dữ liệu"}, {"DATA_TYPE", "Ngân hàng/misa"}}),
    #"Removed Columns2" = Table.RemoveColumns(#"Renamed Columns1",{"Số dư", "Số dư hiện tại", "Số dư giao dịch", "Dư nợ", "Tên nhân viên", "Ngày hóa đơn", "Số hóa đơn", "Dư nợ ngoại tệ", "Dư có ngoại tệ", "Nghiệp vụ", "Loại chứng từ", "Mã nhân viên", "Tên thống kê", "CP hợp lý/không hợp lý", "Kỳ xuất báo cáo", "Số dư cuối kỳ", "Ngày giờ giao dịch", "Tên mục thu/chi", "Tên đối tượng_Misa", "Tên khoản mục CP", "Tên đơn vị_Misa"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns2",{"Loại dữ liệu", "Ngân hàng/misa", "Số tài khoản NH", "Tên file tổng ngân hàng", "Tên file nguồn", "Ngày hạch toán", "Số chứng từ (STT bank)", "Diễn giải", "Số dư đầu kỳ", "Phát sinh tăng tiền", "Phát sinh giảm tiền", "Chi nhánh_Misa", "Hình thức", "MA_DON_VI", "COUNT_MA_DON_VI", "Số phiếu chi", "Mã mục thu/chi", "Mã AP", "Diễn giải riêng", "Tài khoản hạch toán", "TK đối ứng", "Mã đối tượng_Misa", "GNT ngân hàng", "Ghi chú", "CN được chi hộ", "Mã khoản mục CP", "Hợp đồng mua", "Mã thống kê", "Mã đơn vị_Misa"}),
    #"Renamed Columns2" = Table.RenameColumns(#"Reordered Columns",{{"Hình thức", "Hình thức_Saoke"}, {"MA_DON_VI", "Mã ĐV_Saoke"}, {"COUNT_MA_DON_VI", "Count_Mã ĐV_Saoke"}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Renamed Columns2",{{"Số chứng từ (STT bank)", type text}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type1", "Số CT và Diễn giải", each [#"Số chứng từ (STT bank)"]&" --- "&[Diễn giải])
in
    #"Added Custom"