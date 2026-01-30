# Excel to SQL Generator

一个简单的Web应用，用于将Excel文件中的数据转换为SQL语句。

## 功能特点

- 📊 支持 .xlsx 和 .xls 格式的Excel文件
- 🎯 自动识别SQL模板中的参数
- 🔄 自动替换参数并生成SQL语句
- 💾 自动下载生成的SQL文件
- 👀 提供SQL预览功能
- 🎨 现代化的用户界面

## 使用方法

1. 打开 `index.html` 文件
2. 上传Excel文件（点击或拖拽）
3. 选择要处理的Sheet表单
4. 点击"生成SQL文件"按钮
5. 文件将自动下载

## Excel文件格式要求

Excel文件必须按以下格式组织：

| 行号 | 内容 | 示例 |
|------|------|------|
| 第1行 | INSERT SQL示例 | `INSERT INTO users (id, name, email) VALUES (1, '张三', 'test@example.com');` |
| 第2行 | 表头描述 | `id` \| `name` \| `email` |
| 第3行 | 表头（字段名） | `id` \| `name` \| `email` |
| 第4行及以后 | 数据行 | `1` \| `张三` \| `zhangsan@example.com` |

### 工作原理

1. 程序会自动解析第一行的INSERT SQL语句，提取表名和字段列表
2. 根据SQL中的字段名，自动匹配第三行表头中的对应列
3. 将每行数据按照字段顺序填充到SQL语句中
4. 支持大小写不敏感的字段匹配

### 示例

**Excel内容：**

```
第1行: INSERT INTO users (id, name, email) VALUES (1, '示例', 'example@test.com');
第2行: id | 昵称 | 邮箱
第3行: id | name | email
第3行: 1 | 张三 | zhangsan@example.com
第4行: 2 | 李四 | lisi@example.com
```

**生成的SQL：**

```sql
INSERT INTO users (id, name, email) VALUES (1, '张三', 'zhangsan@example.com');

INSERT INTO users (id, name, email) VALUES (2, '李四', 'lisi@example.com');
```

**注意：** 第一行的SQL示例中的具体值（如 `1`, `'示例'`, `'example@test.com'`）不重要，程序只会提取表名和字段名结构。

## 输出文件命名

生成的SQL文件命名格式：`原Excel名+Sheet表单名+日期.sql`

例如：`用户数据+Sheet1+20260123.sql`

## 数据类型处理

- **数字**：直接输出（如：`123`）
- **字符串**：自动添加单引号并转义（如：`'张三'`）
- **空值**：转换为 `NULL`
- **单引号**：自动转义为 `''`

## 技术栈

- HTML5
- CSS3
- JavaScript (ES6+)
- SheetJS (xlsx) - Excel文件解析

## 浏览器兼容性

支持所有现代浏览器：
- Chrome
- Firefox
- Safari
- Edge

## 注意事项

1. 所有处理都在浏览器本地完成，不会上传数据到服务器
2. 确保Excel文件格式正确，否则可能生成错误的SQL
3. 建议在执行SQL前先预览生成的语句
4. 对于大量数据，生成可能需要几秒钟时间

## 许可证

MIT License