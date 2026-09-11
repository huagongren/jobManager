# macOS 云端打包

1. 在 GitHub 创建一个空仓库。无需上传个人简历、成绩单或数据库。
2. 将上传包解压，将其中所有文件（包含 `.github` 和 `scripts` 文件夹）上传到仓库根目录，提交到 main 分支。不要只上传 ZIP 本身。
3. 打开仓库 Actions → Build macOS App。首次提交会自动运行，也可点击 Run workflow 手动运行。
4. 等待两个构建任务成功，在该次运行底部 Artifacts 下载对应版本：
   - arm64：Apple 芯片 Mac（M 系列）。
   - x86_64：Intel Mac。
5. 解压下载包及其中的应用 ZIP，将“秋招管理助手.app”拖入“应用程序”后打开。

不需要在使用者电脑上安装 Python 或 openpyxl。构建在 macOS 15 上进行，旧系统兼容性尚未验证。
应用未进行 Apple Developer 身份签名及公证，macOS 可能阻止首次打开；可在确认来源后按照系统“隐私与安全性”提示允许打开，不需要关闭系统安全保护。

数据保存在当前用户主目录下 `.campus_job_tracker`，复制应用不会自动携带旧电脑的数据。
目前只有云端打包配置，尚未执行 GitHub 构建。工作流会执行启动与数据库初始化检查，不代表全部界面功能已在真实 Mac 上验收。
私有仓库构建可能消耗账户 Actions 额度，请在 GitHub 账户中查看可用额度。
