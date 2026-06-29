# English Teaching Dashboard

Dashboard Streamlit để quản lý bài học tiếng Anh, roadmap theo topic A/B/C, tạo message gửi bạn học, kiểm tra trùng từ vựng ngày trước và phát âm Anh-Anh bằng trình duyệt.

## Cách chạy

1. Giải nén folder.
2. Mở Terminal/CMD tại folder này.
3. Cài thư viện:

```bash
pip install -r requirements.txt
```

4. Chạy app:

```bash
streamlit run app.py
```

5. Trình duyệt sẽ mở dashboard. Nếu không tự mở, copy link local mà Streamlit hiện ra.

## Cấu trúc dữ liệu

File dữ liệu nằm ở:

```text
data/English.xlsx
```

Các sheet chính:

- `Content`: dữ liệu roadmap bài học, dùng để tạo tab A/B/C trên Streamlit.
- `Daily_Lessons`: bài học hằng ngày.
- `Vocabulary_Log`: log từ vựng để kiểm tra trùng.
- `Settings`: ghi chú cấu hình.

## Lưu ý phát âm

Nút phát âm dùng Web Speech API của trình duyệt. App ưu tiên giọng `en-GB` và giọng nam nếu máy/trình duyệt có hỗ trợ. Nếu máy không có giọng nam Anh-Anh, app sẽ fallback sang giọng Anh-Anh bất kỳ hoặc giọng tiếng Anh có sẵn.
