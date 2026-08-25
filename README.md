# SEND-EMAIL 邮件自动发送工具

一个 Windows 图形界面工具，可根据 Excel 名单查找同名工资明细附件，并通过 QQ 邮箱批量发送。

## 功能

- 从 `.xls` 或 `.xlsx` 第一张工作表读取姓名和收件邮箱。
- 在指定文件夹中查找与姓名同名的 `.xlsx` 附件。
- 每发送 18 封邮件自动切换一次发件账号。
- 显示发送进度和成功、失败日志。
- 在运行目录中生成 `失败名单.txt`。

## 数据格式

收件人 Excel 的第一张工作表不需要表头，每行格式如下：

| A 列（姓名/编号） | B 列（邮箱地址） |
| --- | --- |
| 张三 | example@qq.com |

附件文件名必须与 A 列完全对应，例如 `张三.xlsx`。

## 设置邮箱授权码

程序不会在源码或 EXE 中保存密码。运行前，请在 Windows PowerShell 中设置 QQ 邮箱 SMTP 授权码：

```powershell
[Environment]::SetEnvironmentVariable("EMAIL1_PASSWORD", "第一个QQ邮箱授权码", "User")
[Environment]::SetEnvironmentVariable("EMAIL2_PASSWORD", "第二个QQ邮箱授权码", "User")
```

设置后需要重新打开程序。如果要更换默认发件邮箱，也可以设置：

```powershell
[Environment]::SetEnvironmentVariable("EMAIL1_ADDRESS", "第一个发件邮箱", "User")
[Environment]::SetEnvironmentVariable("EMAIL2_ADDRESS", "第二个发件邮箱", "User")
```

请使用 QQ 邮箱设置页面生成的 SMTP 授权码，不要使用邮箱登录密码。

## 使用 EXE

1. 从项目的 `dist` 目录或 GitHub Release 下载 `SEND-EMAIL.exe`。
2. 按上文设置邮箱地址和授权码。
3. 双击运行 EXE。
4. 选择收件人 Excel 和附件文件夹。
5. 点击“开始发送”。

发送真实邮件前，建议先使用少量测试数据验证收件人和附件是否匹配。

## 从源码运行

需要 Python 3.9 或更高版本：

```powershell
python -m pip install -r requirements.txt
python V1.1.py
```

## 打包 EXE

```powershell
python -m pip install pyinstaller
pyinstaller --noconfirm --clean --onefile --windowed --name SEND-EMAIL V1.1.py
```

生成文件位于 `dist/SEND-EMAIL.exe`。
