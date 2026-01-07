# Access到MySQL数据同步工具

## 功能说明

这是一个Windows控制台应用程序，用于定期从两个Access数据库（Data.accdb和NGData.accdb）中提取数据，并将参数数据同步到MySQL数据库。

### 主要功能

1. **连接两个Access数据库**
   - 合格数据：Data.accdb
   - 不合格数据：NGData.accdb

2. **增量数据同步**
   - 根据Time_index字段进行增量查询
   - 自动合并两个数据库的数据
   - 按Time_index排序处理

3. **参数提取和插入**
   - 提取10个参数（空载电流、各种电阻值、泄露等）
   - 从Matrix_Code字段提取partNumber
   - 插入到MySQL的vw.parameters表

4. **定时执行**
   - 默认每5分钟执行一次（可在配置文件中修改）
   - 程序启动时立即执行一次

5. **防重复机制**
   - 使用last_index.txt文件存储最后处理的Time_index
   - 确保数据不重复、不遗漏

## 系统要求

- Windows操作系统
- .NET 6.0 Runtime（如果使用非自包含版本）
- Microsoft Access Database Engine（用于连接.accdb文件）
  - 下载地址：https://www.microsoft.com/zh-cn/download/details.aspx?id=54920
- MySQL数据库服务器

## 安装和配置

### 1. 部署文件

将`publish`目录下的所有文件复制到目标目录，包括：
- `DataSyncTool.exe` - 主程序
- `appsettings.json` - 配置文件
- 其他依赖的DLL文件

### 2. 配置MySQL连接

编辑`appsettings.json`文件，修改MySQL连接信息：

```json
"MySql": {
  "Host": "localhost",        // MySQL服务器地址
  "Port": 3306,               // MySQL端口
  "Database": "vw",           // 数据库名
  "Username": "root",         // 用户名
  "Password": "your_password" // 密码
}
```

### 3. 配置Access数据库路径

编辑`appsettings.json`文件，设置Access数据库文件的路径：

```json
"Access": {
  "DataPath": "Data.accdb",        // 合格数据库路径（相对路径或绝对路径）
  "NGDataPath": "NGData.accdb"     // 不合格数据库路径（相对路径或绝对路径）
}
```

### 4. 配置同步间隔

修改同步执行间隔（单位：分钟）：

```json
"Sync": {
  "IntervalMinutes": 5,           // 执行间隔（分钟）
  "TimeIndexFile": "last_index.txt"  // Time_index存储文件
}
```

### 5. 配置固定字段值

根据实际需求修改固定字段值：

```json
"FixedFields": {
  "Supplier": "M2G",
  "SupplierName": "柯比斯（沈阳）汽车零部件有限公司",
  "VehicleType": "SagitarPA|JettaVS5|JettaVS7",
  "PartName": "燃油泵总成",
  "Station": "Final Test",
  "BTV": "王方权"
}
```

## 使用方法

### 运行程序

1. 确保配置文件`appsettings.json`已正确配置
2. 确保Access数据库文件存在且可访问
3. 确保MySQL数据库服务器可连接
4. 双击运行`DataSyncTool.exe`

### 程序运行说明

- 程序启动时会显示连接状态和配置信息
- 程序会立即执行一次数据同步
- 之后每5分钟（或配置的间隔）自动执行一次
- 按Ctrl+C或关闭窗口可退出程序

### 日志输出

程序会在控制台输出以下信息：
- 配置加载状态
- 数据库连接测试结果
- 每次同步的详细信息：
  - 查询到的记录数
  - 插入的参数数量
  - 处理耗时
  - 错误信息（如有）

## 参数映射

程序会从Access数据库的Data表中提取以下10个参数：

| 参数名称 | Access字段 | 下限 | 上限 |
|---------|-----------|------|------|
| 空载电流 | Current | 0.3 | 1.9 |
| True_Empty电阻 | True_Empty | 287.7 | 292.3 |
| Empty电阻 | Name11Value | 267.7 | 272.3 |
| Reserve电阻 | Reserve | 237.7 | 242.3 |
| 1/4电阻 | 1_4 | 213.7 | 218.3 |
| 1/2电阻 | 1_2 | 159.2 | 162.8 |
| 3/4电阻 | 3_4 | 105.5 | 108.5 |
| Full电阻 | Full | 68.5 | 71.5 |
| True_Full电阻 | True_Full | 50.5 | 53.5 |
| 泄露 | Supply_Leak | -0.14 | 0.35 |

每条Access记录会生成10条MySQL记录（对应10个参数）。

## 故障排除

### 1. 无法连接Access数据库

- 检查Access数据库文件路径是否正确
- 确保已安装Microsoft Access Database Engine
- 检查文件是否被其他程序占用

### 2. 无法连接MySQL数据库

- 检查MySQL服务器是否运行
- 检查连接信息（主机、端口、用户名、密码）是否正确
- 检查防火墙设置
- 检查MySQL用户权限

### 3. 数据未同步

- 检查Time_index是否正确递增
- 查看控制台输出的错误信息
- 检查last_index.txt文件内容

### 4. 程序异常退出

- 查看控制台输出的详细错误信息
- 检查配置文件格式是否正确（JSON格式）
- 确保所有必需的字段都已配置

## 注意事项

1. **首次运行**：如果last_index.txt不存在，程序会从Time_index=0开始查询所有数据
2. **数据安全**：程序使用事务确保数据一致性，如果插入失败会自动回滚
3. **性能考虑**：如果数据量很大，建议适当调整同步间隔
4. **文件权限**：确保程序有权限读取Access数据库文件和写入last_index.txt文件

## 技术支持

如有问题，请检查：
1. 配置文件格式是否正确
2. 数据库连接是否正常
3. 控制台输出的错误信息

