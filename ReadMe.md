# Task字段定义表

| 字段名称 | 数据类型 | 备注 |
| :--- | :--- | :--- |
| **ID** | 字符串 (string) | 唯一标识符（例如：任务ID，记录ID） |
| **ChipNumber** | 字符串 (string) | 芯片编号 |
| **TaskName** | 字符串 (string) | 任务名称 |
| **Status** | 字符串 (string) | 当前状态（例如：已完成, 运行中, 失败） |
| **Level** | 字符串 (string) | 优先级或级别（例如：高, 中, 低） |
| **Major** | 字符串 (string) | 设备的主型号 |
| **Minor** | 时间类型 (string) | 设备的子型号， 可以为空 |
| **StartDate** | 时间类型 (DateTime) | 任务/操作的开始日期和时间 |
| **EndDate** | 时间类型 (DateTime) | 任务/操作的结束日期和时间 |
| **DataStatus** | 字符串 (string) | 数据处理状态 |
| **FilesStatus** | 字符串 (string) | 相关文件（或文件组）的状态 |
| **Contions** | 字符串 (string) | 任务/操作的条件, 用json 表示 |