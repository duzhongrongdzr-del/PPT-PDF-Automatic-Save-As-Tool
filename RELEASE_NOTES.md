Windows x64 0.2.0 测试版，新增 **PDF 打开后自动留存**。

下载 **ZhitingPPTSave-Windows-x64.zip**，解压整个文件夹，运行 **ZhitingPPTSave.exe**。无需安装 Python。

- PPT / PDF 共用原有保存目录，PDF 功能默认启用。
- 保留原文件名、按内容去重、同名不同内容分目录保存。
- PDF 识别尝试覆盖 WPS PDF、Adobe Reader / Acrobat、福昕、SumatraPDF、PDF-XChange，以及浏览器中的本地 PDF。
- 没有本地原文件的在线 PDF 暂不支持自动下载。具体阅读器版本、浏览器新标签页和元数据标题等情况仍需实测。

更新前先右键旧程序托盘图标 → 退出，再运行新版本。原有保存目录会保留；如果更换程序位置，取消后重新勾选开机启动。

发布前运行归档和路径识别测试、Windows 集成检查、真实 Edge 打开本地 PDF 的识别与复制检查，以及打包后的程序冒烟测试。以上检查不代表所有 WPS / Adobe 版本已经通过实机验收。

文件仅在本机处理，不上传。未做商业代码签名，SHA256SUMS.txt 提供压缩包校验值。
