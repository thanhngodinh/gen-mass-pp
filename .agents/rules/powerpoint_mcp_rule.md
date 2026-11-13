# Bật PowerPoint MCP

Kiểm tra registry phiên hiện tại có `ppt_*` và `church_*`. Nếu chưa có, kiểm tra source `tools/ppt-mcp`, README và cấu hình `%USERPROFILE%\.codex\config.toml`:

```toml
[mcp_servers.powerpoint]
command = "uv"
args = ["--directory", "C:\\PP\\tools\\ppt-mcp", "run", "mcp", "run", "src/server.py"]
```

Nếu cần, dùng đường dẫn thực tế tới `uv.exe` và cài dependencies bằng `uv sync` trong `tools/ppt-mcp`. Bật hoặc sửa cấu hình trước, yêu cầu reset Codex, rồi xác nhận tool xuất hiện và gọi tool chỉ đọc. Nếu chưa thể reset, dừng thao tác PowerPoint và báo người dùng. Mọi thao tác với PowerPoint dùng MCP; PDF nhạc đọc bằng vision riêng.
