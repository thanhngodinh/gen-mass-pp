# PowerPoint Thánh Lễ — hướng dẫn sử dụng

Dự án giúp Codex tạo, sửa và kiểm tra PowerPoint Thánh Lễ theo phôi Giáo xứ Tân Định. AI đọc lịch hát và sheet nhạc, tra bài đọc rồi dùng PowerPoint MCP để thao tác trên Microsoft PowerPoint. Quy tắc chi tiết dành cho AI nằm trong `.agents/skills/church-presentation-generator/SKILL.md`; `AGENTS.md` yêu cầu AI đọc skill và dùng MCP.

## Chuẩn bị máy Windows

1. Cài Microsoft PowerPoint bản desktop và mở PowerPoint ít nhất một lần. Cài Codex desktop và đăng nhập tài khoản có quyền sử dụng.
2. Cài `uv` theo https://docs.astral.sh/uv/getting-started/installation/; mở lại PowerShell và kiểm tra `uv --version`.
3. Tải hoặc clone repo, ví dụ vào `C:\PP`. Trong PowerShell chạy:

```powershell
cd C:\PP\tools\ppt-mcp
uv sync
```

Lệnh này tải Python và thư viện cần thiết nếu máy chưa có. Cần Internet khi cài phụ thuộc lần đầu.

## Kết nối MCP với Codex

Mở hoặc tạo `%USERPROFILE%\.codex\config.toml` và thêm cấu hình sau. Nếu thư mục repo ở nơi khác, đổi đường dẫn trong `args`; trong TOML, mỗi dấu `\` của đường dẫn Windows được viết thành `\\`.

```toml
[mcp_servers.powerpoint]
command = "uv"
args = ["--directory", "C:\\PP\\tools\\ppt-mcp", "run", "mcp", "run", "src/server.py"]
```

Giữ nguyên các cấu hình khác đang có; nếu đã có mục `mcp_servers.powerpoint`, sửa mục đó thay vì thêm trùng. Nếu Codex không tìm thấy `uv`, thay `command` bằng đường dẫn đầy đủ của `uv.exe`. Khởi động lại Codex, mở repo và yêu cầu: “Kiểm tra PowerPoint MCP đã sẵn sàng chưa”. AI phải thấy các tool `ppt_*` và `church_*` và gọi một tool chỉ đọc. Nếu chưa thấy, kiểm tra lại đường dẫn, `uv sync`, cấu hình rồi khởi động lại Codex. Chi tiết: `.agents/rules/powerpoint_mcp_rule.md`.

## Làm PowerPoint mỗi tuần

1. Tải PDF sheet nhạc từ Zalo vào Downloads hoặc gửi file trực tiếp cho Codex. Có thể gửi thêm đường dẫn phôi PowerPoint cần dùng.
2. Dán lịch hát, ngày lễ, tuần và thứ tự hát vào chat, ví dụ:

```text
Tuần 33, lễ Thứ Hai
NL: Nơi Nhà Chúa (1-ĐK-ĐK)
DL: Lời Dâng (1-ĐK-ĐK)
RL: Lắng Nghe Lời Chúa
TL: Từ Chốn Luyện Hình (1-ĐK-ĐK)
Làm PP dựa trên file phôi tôi gửi.
```

3. AI đối chiếu file với tuần và lịch, tạo `week/<tuần>/in/`, sao chép tài liệu đầu vào và ghi lịch vào `calendar.md`. Nếu có nhiều file không phân biệt được, AI sẽ hỏi lại. Đầu ra nằm ở `week/<tuần>/out/`; file gốc được giữ lại.
4. AI đọc sheet bằng vision, tra bài đọc tại https://kpv.vn/liturgy/mass-readings, chỉnh bản sao phôi bằng MCP và kiểm tra nội dung cùng ảnh xem trước. Phôi mặc định nằm ở `resources/templates/`.

Khi bắt đầu lượt mới, nói rõ tuần/ngày và yêu cầu. AI phải kiểm tra MCP đã nạp trước khi dùng; MCP thường được Codex khởi chạy từ cấu hình, không cần tự chạy lệnh server cho từng slide.

## Kiểm tra PowerPoint đã có

Gửi file hoặc đường dẫn và nhắn “Kiểm tra nội dung PP này theo lịch tuần …”. AI mở file chỉ đọc qua MCP, đối chiếu ngày lễ, bài đọc, bài hát và trình tự, sau đó báo điểm cần xem lại theo số slide. Muốn AI sửa, hãy yêu cầu sửa riêng.

## Các quy tắc trình chiếu chính

Giữ ảnh, font, cỡ chữ và bố cục phôi; được thêm hoặc xóa slide để lời ca vừa khung. Tiêu đề lễ kính/trọng thể ưu tiên tên lễ; tiêu đề bìa và divider dùng 24 pt, viết TN cho Thường Niên. Thứ Bảy có thêm bài Hiệp Lễ cố định “Các Con Là Ánh Sáng”. Giữ hai section Kinh Tin Kính để người trình chiếu chọn; thường dùng Nicea.

Thư mục `week/` và `scratch/` là dữ liệu làm việc trên máy, không đưa lên Git. Không gửi tài khoản, mật khẩu hoặc token qua chat.
