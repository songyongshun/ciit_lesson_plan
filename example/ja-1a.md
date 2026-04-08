# 项目名称
Linux 基本配置

## 授课类型
理实一体化

## 授课周次
第5周

## 授课学时
2学时

## 教学目标
### 知识目标
- 掌握用户账户管理、时区与本地化（locale）的基本概念与实现方法。
- 理解包管理器（APT/DPKG）与软件仓库的工作机制与安全原理。
- 了解系统服务管理与软件更新策略对稳定性的影响。

## 能力目标
- 能配置用户账户与 sudo 权限，并完成时区与本地化设置。
- 能使用包管理工具完成软件安装、更新、移除与故障排查。
- 能记录并复现配置步骤，形成标准化配置文档。

## 素质目标
- 培养严格验证配置变更与回滚的习惯。
- 强化安全意识，避免使用 root 进行日常操作。
- 提高日志记录与变更审计的意识。

## 学情分析
多数学生完成了安装或已搭建 WSL，具备基础操作能力，可直接进入配置练习。

## 教学重点
- 包管理工具的使用与系统更新流程。

## 教学难点
- 理解本地化（locale）设置对系统行为的影响。

## 教学方法
演示＋练习＋分组讨论

## 教材资源
发行版官方文档、包管理器手册（APT）

## 课前:教学内容
- 阅读关于 `apt` 的入门文档。

## 课前:学生活动
- 准备可用的测试账户并记录管理员密码。

## 课前:教师活动
- 准备演示脚本与常用软件列表。

## 项目导入
### 项目导入:教学内容
示例演示如何通过包管理安装 `htop`、`git` 并执行系统更新。

### 项目导入:学生活动
- 小组内讨论常用软件并列出安装顺序。

### 项目导入:教师活动
- 演示 `sudo apt update && sudo apt upgrade` 与 `apt install` 的差别。

## 内容展开:教学内容
![系统配置示意](images/lesson-02-diagram.png)

本节通过理论与大量实操相结合的方式，系统性讲解 Linux 基本配置的常见任务与最佳实践，重点覆盖用户与权限管理、时区与本地化设置、包管理与软件仓库、系统服务管理以及常见故障排查与备份恢复策略。目标是让学生能够独立完成从初始配置到常规维护的一套标准化流程，并具备在出现配置失败时的快速定位与回滚能力。

一、用户与权限管理（约700字）
1. 概念梳理：说明 Linux 的多用户模型、UID/GID 的含义、超级用户（root）与普通用户的区别。强调安全原则：日常操作使用普通账户并通过 `sudo` 临时获取管理员权限，避免长期以 root 身份登录导致的风险。
2. 常用命令与操作：
- 新增用户：`sudo adduser <username>`（会交互创建家目录和默认 shell）。
- 删除用户：`sudo deluser <username>`，若需删除家目录：`sudo deluser --remove-home <username>`。
- 修改用户组：`sudo usermod -aG <group> <username>`（注意 `-a` 参数用于追加）。
- 查看用户信息：`id <username>`、`getent passwd <username>`、`groups <username>`。
3. sudo 管理：
- `sudo` 的工作机制（通过 `/etc/sudoers` 定义权限），教会学生使用 `sudo visudo` 安全编辑 sudoers 文件，并举例如何为某个用户授予部分命令权限（使用别名与限定主机）。
4. 实操练习：
- 任务：创建用户 `student1`，将其加入 `sudo` 组，限制其只能使用 `systemctl` 与 `apt` 相关命令（示例 sudoers 条目）。
- 验证：以 `student1` 登录并尝试执行被允许/被拒绝的命令，记录日志并说明发生原因。
5. 常见问题与排查：
- 无法使用 sudo：检查用户是否在 sudo 组、`/etc/sudoers` 是否误改。查看 `/var/log/auth.log` 中的相关条目。

二、时区与本地化（locale）设置（约400字）
1. 核心概念：解释时区（timezone）与本地化（locale）的区别与联系，说明字符编码（UTF-8）对文件名、终端显示与文本处理的重要性。
2. 常用命令：
- 列出时区：`timedatectl list-timezones`
- 设置时区：`sudo timedatectl set-timezone Asia/Shanghai`
- 查看当前时区与 NTP 状态：`timedatectl status`
- 查看 locale：`locale`，生成 locale：`sudo locale-gen zh_CN.UTF-8`，更新默认：`sudo update-locale LANG=zh_CN.UTF-8`
3. 实操练习：
- 任务：将系统时区设置为 `Asia/Shanghai`，确保 `date` 命令输出为本地时间；设置并验证 `zh_CN.UTF-8` locale，使用 `locale charmap` 验证编码为 UTF-8。
4. 常见陷阱：
- WSL 与容器环境中时区/locale 的特殊处理（有时需要在宿主机或容器启动脚本中设置），字符编码不一致会导致日志或文件名显示异常。

三、包管理与软件仓库（约900字）
1. 基本概念：介绍包管理器的职责：软件安装、依赖解析、版本管理与安全校验。以 Debian/Ubuntu 的 APT 生态为主线，简要提及其他发行版（如 CentOS 的 yum/dnf、Arch 的 pacman）区别。
2. 核心命令与流程：
- 更新索引：`sudo apt update`（与 `apt-get update` 的等价关系）。
- 升级系统：`sudo apt upgrade`（保守升级）与 `sudo apt full-upgrade`（可变动依赖）。
- 安装软件：`sudo apt install <package>`；卸载软件：`sudo apt remove <package>`；彻底删除（含配置）：`sudo apt purge <package>`。
- 清理：`sudo apt autoremove`（删除不再需要的依赖）、`sudo apt autoclean`/`sudo apt clean`（清理索引缓存）。
3. 软件仓库管理：
- `/etc/apt/sources.list` 与 `/etc/apt/sources.list.d/` 的结构与优先级。
- 添加第三方仓库：`sudo add-apt-repository ppa:...` 或手动添加 `.list` 文件，注意签名（APT 使用 GPG 验证包仓库）：`apt-key` 的衰退与 `signed-by` 机制。
4. 常见故障与排查：
- 索引更新失败：检查网络、`/etc/hosts` 与 DNS、代理设置；查看 `sudo apt update` 的错误信息并定位具体源。
- 发生依赖冲突或损坏包：使用 `sudo apt --fix-broken install`、`sudo dpkg --configure -a`。查看 `/var/log/apt/term.log` 与 `/var/log/dpkg.log` 获取详细错误。
- 软件版本回滚策略：保留配置备份、使用 `apt-mark hold <package>` 锁定版本或从 `/var/cache/apt/archives/` 手动重装旧包。
5. 示例演示：
- 教师演示安装 `git`、`htop`，演示如何查看包信息：`apt show <package>`、`apt-cache policy <package>`，并模拟网络中断查错流程。

四、系统服务管理与日志（约400字）
1. systemd 与 systemctl 简介：介绍 systemd 的概念、服务单元（unit）类型（service、socket、timer、target）与常用管理命令。
2. 常用操作：
- 启动/停止/重启：`sudo systemctl start/stop/restart <service>`
- 设置开机自启：`sudo systemctl enable <service>`，取消自启：`sudo systemctl disable <service>`
- 查询状态：`sudo systemctl status <service>`，查看日志：`sudo journalctl -u <service> -f`（实时跟踪）。
3. 实操练习：
- 安装并启用 `ssh`（演示如何查看端口与防火墙状态），要求学生完成服务的启停与日志查看，并记录启动失败时的诊断步骤。

五、备份、变更管理与回滚（约200字）
1. 配置文件备份策略：在修改 `/etc` 下关键文件前，使用版本控制（如在 `/etc` 下建立 git 仓库备份）、或保存带时间戳的副本：`sudo cp /etc/apt/sources.list /etc/apt/sources.list.bak.$(date +%F-%T)`。
2. 变更记录与自动化：鼓励使用脚本化配置或 Ansible 一类的工具来保持可重复性，并在遇到故障时快速回滚。

六、课堂练习与评估（约200字）
1. 分组实操任务（35分钟）：每组需完成用户创建与权限配置、时区/locale 设置、安装 `git` 与 `htop`、启用 `ssh` 并记录全套命令与输出。教师巡回检查并收集配置日志。
2. 问答检查（10分钟）：要求每位学生回答三题：解释 `sudo` 与 `su` 的区别、列举三条 apt 常用命令并说明区别、说明遇到包依赖冲突的排查步骤。

七、延伸阅读与作业指引
- 推荐官方文档与若干优秀博客案例，布置课后作业（详见“课后:作业”）。
## 内容展开:学生活动
- 按步骤创建用户并配置 sudo 权限；设置时区并验证；安装 `git` 并克隆一个仓库。

## 内容展开:教师活动
- 指导并检查学生操作，演示故障排查技巧。

## 课后:作业
- 完成配置日志并在课堂系统中提交，描述遇到的问题与解决方法。

## 教学反思
### 教学反思:教学效果
（在此记录课堂效果、学生掌握情况与教学中观察到的问题。）

### 教学反思:诊断
（分析教学中发现的知识薄弱环节与原因。）

### 教学反思:改进
（下一步改进策略与建议。）
