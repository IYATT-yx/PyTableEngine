# extensions 插件目录

请将插件文件夹放在此目录下，结构如图示：  

```text
extensions/
│
├── README.md                          # 插件目录说明文档（当前文件）
│
├── IYATT_BatchMergeByColumns/         # 示例插件 1：按列批量合并
│   ├── __init__.py                    # (可选) 模块标识
│   ├── IYATT_BatchMergeByColumns.py   # 插件主入口逻辑文件
│   └── requirements.txt               # (可选) 插件专属依赖清单
│
├── IYATT_CopyMergedByMatrix/          # 示例插件 2：跨合并单元格矩阵复制
│   ├── IYATT_CopyMergedByMatrix.py
│   └── requirements.txt
│
└── IYATT_DemoPlugin/                  # 示例插件 3：模板插件
    └── IYATT_DemoPlugin.py
```