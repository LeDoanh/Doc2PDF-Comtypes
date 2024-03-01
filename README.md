Source test dùng lib comtypes để convert word qua pdf.
Note: Lần đầu chạy hơi lâu

** Hướng dẫn sử dụng:

1. Tạo evn riêng cho source:
    - Tạo evn ảo (Lần đầu chạy source): 
        python -m venv doc2pdf

    - Activate evn (Mỗi lần mở source): 
        doc2pdf\Scripts\activate

    - Install package từ file requirements (Lần đầu chạy source)
        pip install -r requirements.txt

2. Chạy source:
    - Lấy đường đẫn file cần convert:
        PathIn = FileIn.doc / FileIn.docx

    - Chọn đường đẫn file lưu
        PathOut = FileOut.pdf

    - Chạy lệnh terminal hoặc cmd (Đổi PathIn, PathOut )
        python doc2pdf.py "PathIn" "PathOut"
