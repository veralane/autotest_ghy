# 报告生成框架 - 完整链路说明

## 整体架构

```
┌──────────┐    ┌──────────────┐    ┌──────────────┐    ┌──────────────┐    ┌──────────────┐
│  MF4文件  │ →  │  信号提取     │ →  │  事件检测     │ →  │  指标计算     │ →  │  报告生成     │
│  (原始数据) │    │  (asammdf)   │    │  (制动事件)   │    │  (7项指标)    │    │  (Word .docx) │
└──────────┘    └──────────────┘    └──────────────┘    └──────────────┘    └──────────────┘
                      ↑                    ↑                    ↑                    ↑
                signal_mapping.json   event_detection      计算公式            report_data.json
                (配置信号名)          (配置触发条件)       (computed_signals)   (中间数据)
```

## 完整流程（5步）

### 步骤1：配置信号映射

编辑 `config/signal_mapping.json`，将MF4中的实际信号名填入。

**如果你不知道MF4里有哪些信号，先列出：**
```powershell
python scripts/mf4_parser.py --input data.mf4 --list-signals
python scripts/mf4_parser.py --input data.mf4 --list-signals --filter brake
```

**然后填入信号名：**
```json
"common": {
    "VehicleSpeed":  { "signal": "ESP_VehicleSpeed",   ... },
    "VehicleDecel":  { "signal": "ESP_LongDecel",       ... },
    "SteeringAngle": { "signal": "SAS_SteeringAngle",   ... },
    "BrakePedal":    { "signal": "ESP_BrakePedalSt",    ... },
    ...
}
```

### 步骤2：配置事件检测条件

同样在 `signal_mapping.json` 的 `event_detection` 部分：

```json
"event_detection": {
    "brake_trigger_signal": "ESP_BrakePedalSt",
    "brake_trigger_threshold": 0.5,
    "speed_signal": "ESP_VehicleSpeed",
    "speed_bands": [
        { "label": "50",  "min": 45, "max": 55 },
        { "label": "80",  "min": 75, "max": 85 },
        { "label": "100", "min": 95, "max": 105 }
    ]
}
```

### 步骤3：解析MF4文件

```powershell
# 完整解析
python scripts/mf4_parser.py --input data.mf4 --config config/signal_mapping.json --output output/parsed_data.json

# 或仅解析不生成报告
.\scripts\run.ps1 -Mf4File data.mf4 -ParseOnly
```

**输出：** `output/parsed_data.json`（中间JSON数据）

### 步骤4：人工补充

MF4解析只能提取客观数据，以下字段需要人工补充：

| 字段 | 说明 | 在哪里填 |
|------|------|---------|
| 主观评分 | 驾驶员主观评价1-10分 | parsed_data.json |
| 主观评价 | 文字描述 | parsed_data.json |
| 测试路面 | 干沥青/湿沥青/对开 | parsed_data.json |
| 测试温度/湿度 | 环境条件 | parsed_data.json |
| 结论 | 通过/失败 | parsed_data.json |
| 项目基本信息 | 项目名、版本号等 | parsed_data.json |

### 步骤5：生成Word报告

```powershell
# 从解析后的JSON生成报告
.\scripts\run.ps1 -DataFile output/parsed_data.json -SkipParse

# 或一键完成（MF4 → 报告）
.\scripts\run.ps1 -Mf4File data.mf4
```

## 文件结构

```
vehicle-test-report/
├── config/
│   ├── signal_mapping.json           ← 🔧 步骤1&2：配置信号映射和事件检测
│   ├── report_data_template.json     ← 📝 数据模板（手动填写时用）
│   └── test_requirements.json        ← 📋 测试要求配置
├── scripts/
│   ├── mf4_parser.py                 ← ⚙️ 步骤3：MF4解析+信号提取+指标计算
│   ├── generate_report.ps1           ← ⚙️ 步骤5：JSON → Word报告
│   └── run.ps1                       ← 🚀 一键运行入口
├── output/                           ← 📄 输出目录
│   ├── parsed_data_20260416_115100.json   ← 中间JSON数据
│   └── vehicle_test_report_20260416_115100.docx  ← 最终报告
└── FRAMEWORK.md                      ← 📖 本文档
```

## 指标计算说明

| 指标 | 计算方法 | 输入信号 |
|------|---------|---------|
| 平均减速度 | 制动区间内减速度绝对值的均值 | VehicleDecel |
| 制动距离 | 对车速在制动区间内积分 | VehicleSpeed |
| 减速度峰谷差值 | 减速度相邻峰谷差的最大值 | VehicleDecel |
| 转向修正角 | 制动区间内方向盘最大偏转角 | SteeringAngle |
| 车轮抱死时间 | 滑移率>0.8的累计时间 | WheelSlip_* |
| 附着系数利用率 | 平均减速度/(μ×9.81)×100% | VehicleDecel + RoadMu |
| 主观评分 | 人工填写 | — |

## 使用场景

### 场景A：有MF4数据文件
```powershell
# 1. 先看看MF4里有什么信号
.\scripts\run.ps1 -Mf4File data.mf4 -ListSignals

# 2. 配置 signal_mapping.json 中的信号名

# 3. 一键生成
.\scripts\run.ps1 -Mf4File data.mf4
```

### 场景B：手动填写数据
```powershell
# 1. 复制模板
copy config\report_data_template.json my_data.json

# 2. 编辑 my_data.json 填入测试数据

# 3. 生成报告
.\scripts\run.ps1 -DataFile my_data.json -SkipParse
```

### 场景C：MF4解析后人工补充
```powershell
# 1. 解析MF4
.\scripts\run.ps1 -Mf4File data.mf4 -ParseOnly

# 2. 编辑 output/parsed_data_*.json 补充主观评分等

# 3. 生成报告
.\scripts\run.ps1 -DataFile output/parsed_data_xxx.json -SkipParse
```

## 环境依赖

| 组件 | 用途 | 安装 |
|------|------|------|
| Python 3.8+ | MF4解析 | python.org |
| asammdf | MF4文件读取 | pip install asammdf |
| numpy | 数值计算 | pip install numpy |
| pandas | 数据处理 | pip install pandas |
| Microsoft Word | 报告生成 | 已安装 |

**快速安装Python依赖：**
```bash
pip install asammdf numpy pandas
```
