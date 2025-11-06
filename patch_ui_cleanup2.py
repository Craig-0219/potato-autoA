from pathlib import Path

path = Path(r"autoa/ui.py")
lines = path.read_text(encoding="utf-8").splitlines()
idx = 0
while idx < len(lines):
    if 'messagebox.showwarning("箭頭校正"' in lines[idx]:
        lines[idx] = '            messagebox.showwarning("箭頭校正", "請先補齊模板檔案:\\n" + "\\n".join(missing))'
        while idx + 1 < len(lines) and lines[idx + 1].strip().startswith('"'):
            lines.pop(idx + 1)
    elif 'messagebox.showinfo("箭頭校正結果"' in lines[idx]:
        lines[idx] = '        messagebox.showinfo("箭頭校正結果", "\\n".join(info_lines))'
        while idx + 1 < len(lines) and lines[idx + 1].strip().startswith('"'):
            lines.pop(idx + 1)
    else:
        idx += 1

path.write_text("\n".join(lines) + "\n", encoding="utf-8")
