# 邮件发送项目使用说明

本项目包含多个邮件发送相关的 Python 脚本，支持使用 SMTP 和 yagmail 发送邮件。

## 项目功能

### 1. 批量发送成绩邮件 (`send_grades.py`)

从 CSV 文件读取联系人信息，批量发送个性化的成绩通知邮件。

#### 功能特点

- 读取 `contacts_file.csv` 文件中的联系人信息
- 使用 HTML 格式发送美观的成绩通知邮件
- 支持个性化的姓名和成绩显示
- 完整的错误处理和发送状态统计
- 支持从环境变量或交互式输入获取密码

#### 使用方法

1. **准备联系人 CSV 文件**

   编辑 `contacts_file.csv` 文件，格式如下：

   ```csv
   name,email,grade
   張小明,student1@gmail.com,85
   李美華,student2@gmail.com,92
   王大偉,student3@gmail.com,78
   ```

2. **配置邮件发送者信息**

   创建 `.env` 文件（在 `src` 目录或项目根目录），添加以下配置：

   ```env
   GMAIL_PASSWORD=your_app_password
   GMAIL_SENDER=your_email@gmail.com
   ```

   **注意**：Gmail 需要使用「应用程式密码」，不是一般密码。取得方法：Google 帐户 → 安全性 → 两步骤验证 → 应用程式密码

3. **执行脚本**

   ```powershell
   cd src
   python send_grades.py
   ```

   脚本会：
   - 自动读取 CSV 文件
   - 显示联系人列表
   - 提示确认后发送邮件
   - 显示发送结果统计

#### 邮件模板示例

脚本会发送包含以下内容的 HTML 邮件：
- 友好的问候语（包含学生姓名）
- 成绩显示（大号字体，突出显示）
- 鼓励性文字

#### CSV 文件格式

- 第一行：标题行 `name,email,grade`
- 数据行：每行包含姓名、邮箱和成绩
- 支持空行（会自动跳过）
- 编码：UTF-8

---

### 2. yagmail 邮件发送 (`testmail.py`)

使用 yagmail 库发送 HTML 邮件（带附件和图片）。

#### 使用方法

```powershell
cd src
python testmail.py
```

#### 功能

- 发送 HTML 格式邮件
- 支持附件（PDF）
- 支持内嵌图片
- 从 `name_list.md` 读取收件人列表

---

## 安装依赖

### 1. 激活虚拟环境

```powershell
# 在项目根目录
.\mail\Scripts\Activate.ps1
```

### 2. 安装依赖套件

```powershell
pip install -r requirements.txt
```

### 3. 依赖列表

- `yagmail==0.15.293` - yagmail 邮件发送库
- `python-dotenv==1.0.0` - 环境变量管理
- `twilio` - 短信发送（如需要）

---

## 环境配置

### 创建 `.env` 文件

在 `src` 目录或项目根目录创建 `.env` 文件：

```env
# Gmail 应用程式密码
GMAIL_PASSWORD=your_app_password_here

# 发送者邮箱（可选，默认使用 peter861005@gmail.com）
GMAIL_SENDER=your_email@gmail.com

# 其他可能的变量名（send_grades.py 会尝试读取）
PASSWORD=your_app_password_here
EMAIL_PASSWORD=your_app_password_here
MAIL_PASSWORD=your_app_password_here
```

### Gmail 应用程式密码设置步骤

1. 登录 Google 帐户
2. 进入「安全性」设置
3. 启用「两步骤验证」
4. 创建「应用程式密码」
5. 选择「邮件」和「其他（自定义名称）」
6. 将生成的密码复制到 `.env` 文件

---

## 文件结构

```
email/
├── src/
│   ├── send_grades.py          # 批量发送成绩邮件脚本
│   ├── testmail.py             # yagmail 测试邮件脚本
│   ├── sendsms.py              # 短信发送脚本（如有）
│   ├── contacts_file.csv       # 联系人 CSV 文件
│   ├── name_list.md            # 收件人列表（testmail.py 使用）
│   ├── 示例.md                 # 邮件发送示例代码
│   ├── requirements.txt        # Python 依赖套件
│   └── README.md               # 本说明文件
├── .env                        # 环境变量配置（需自行创建）
└── mail/                       # Python 虚拟环境
```

---

## 常见问题

### Q: 发送邮件失败，提示认证错误？

**A:** 请确认：
1. 已启用 Gmail 两步骤验证
2. 使用的是「应用程式密码」，不是帐户密码
3. `.env` 文件中的密码格式正确（没有空格或引号）
4. `.env` 文件位置正确（`src` 目录或项目根目录）

### Q: CSV 文件读取失败？

**A:** 请确认：
1. CSV 文件格式正确（包含标题行 `name,email,grade`）
2. 文件编码为 UTF-8
3. 每行数据完整（姓名、邮箱、成绩三个字段）

### Q: 如何修改邮件模板？

**A:** 编辑 `send_grades.py` 中的 `create_email_template()` 函数，修改 HTML 内容和样式。

### Q: 如何添加更多联系人？

**A:** 直接在 `contacts_file.csv` 文件中添加新行即可：
```csv
新联系人,new_email@example.com,90
```

### Q: 能否发送给多个不同的邮箱？

**A:** 可以。在 CSV 文件中，每个联系人的 `email` 字段可以设置为不同的邮箱地址。

---

## 示例

### 示例 1：批量发送成绩邮件

1. 准备 `contacts_file.csv`：
   ```csv
   name,email,grade
   張小明,student1@gmail.com,85
   李美華,student2@gmail.com,92
   ```

2. 配置 `.env`：
   ```env
   GMAIL_PASSWORD=your_app_password
   ```

3. 运行脚本：
   ```powershell
   python send_grades.py
   ```

4. 确认发送后，脚本会自动发送邮件给所有联系人。

---

## 注意事项

1. **安全**：`.env` 文件包含敏感信息，请勿提交到版本控制系统
2. **Gmail 限制**：Gmail 每日发送邮件有数量限制，批量发送时请注意
3. **测试**：建议先用测试邮箱进行测试，确认无误后再发送给实际收件人
4. **编码**：所有文件使用 UTF-8 编码，确保中文正确显示

---

## 开发参考

- `示例.md` - 包含 SMTP、yagmail、Twilio 等各种邮件发送示例代码
- `經驗.md` - 开发经验和注意事项
