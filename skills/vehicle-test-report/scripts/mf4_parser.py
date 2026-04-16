#!/usr/bin/env python3
"""
MF4数据解析器 - 从MF4文件中提取CAN信号并计算测试指标

功能：
1. 读取MF4文件，解析CAN总线数据
2. 根据信号映射配置提取目标信号
3. 自动识别制动事件（基于刹车踏板/减速度触发条件）
4. 计算各测试指标（平均减速度、制动距离、峰谷差值等）
5. 输出标准JSON数据，供报告生成脚本使用

依赖：
    pip install asammdf numpy pandas

Usage:
    python mf4_parser.py --input data.mf4 --config signal_mapping.json --output report_data.json
    python mf4_parser.py --input data.mf4 --auto-detect  # 自动识别制动事件
"""

import argparse
import json
import os
import sys
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Optional, Tuple, Any

import numpy as np
import pandas as pd

try:
    from asammdf import MDF
except ImportError:
    print("请安装 asammdf: pip install asammdf")
    sys.exit(1)


# ============================================================
# 信号提取模块
# ============================================================

class SignalExtractor:
    """从MF4文件中提取CAN信号"""

    def __init__(self, mdf_path: str, signal_config: dict):
        """
        Args:
            mdf_path: MF4文件路径
            signal_config: 信号映射配置（signal_mapping.json内容）
        """
        self.mdf_path = mdf_path
        self.signal_config = signal_config
        self.mdf = None
        self.available_signals = []

    def load(self):
        """加载MF4文件"""
        print(f"加载MF4文件: {self.mdf_path}")
        self.mdf = MDF(self.mdf_path)
        self.available_signals = self.mdf.channels_db
        print(f"  通道数量: {len(self.available_signals)}")
        return self

    def list_signals(self, keyword: str = None) -> List[str]:
        """列出可用信号，支持关键词过滤"""
        if keyword:
            return [s for s in self.available_signals if keyword.lower() in s.lower()]
        return list(self.available_signals)

    def get_signal(self, signal_name: str) -> Optional[pd.Series]:
        """提取单个信号"""
        if signal_name not in self.available_signals:
            print(f"  ⚠ 信号未找到: {signal_name}")
            return None
        data = self.mdf.get(signal_name)
        # 转换为pandas Series
        timestamps = data.timestamps
        values = data.samples
        return pd.Series(values, index=pd.to_datetime(timestamps, unit='s'), name=signal_name)

    def get_signals_batch(self, signal_names: List[str]) -> Dict[str, pd.Series]:
        """批量提取信号"""
        result = {}
        for name in signal_names:
            sig = self.get_signal(name)
            if sig is not None:
                result[name] = sig
        return result

    def extract_mapped_signals(self) -> Dict[str, pd.Series]:
        """根据signal_mapping.json提取所有映射的信号"""
        signals = {}
        mappings = self.signal_config.get("signal_mappings", {})

        for test_type, fields in mappings.items():
            if test_type.startswith("_"):
                continue
            for field_name, field_config in fields.items():
                signal_name = field_config.get("signal")
                if signal_name and signal_name not in signals:
                    sig = self.get_signal(signal_name)
                    if sig is not None:
                        signals[signal_name] = sig
                        print(f"  ✓ 提取信号: {signal_name}")

        return signals

    def close(self):
        """关闭MF4文件"""
        if self.mdf:
            self.mdf.close()


# ============================================================
# 制动事件检测模块
# ============================================================

class BrakeEventDetector:
    """自动检测制动事件"""

    def __init__(self, config: dict = None):
        self.config = config or {
            "brake_trigger_signal": "BrakePedal",     # 刹车踏板信号名
            "brake_trigger_threshold": 0.5,            # 刹车触发阈值
            "decel_trigger_signal": "VehicleDecel",   # 减速度信号名
            "decel_trigger_threshold": -2.0,           # 减速度触发阈值 (m/s²)
            "min_duration_s": 1.0,                     # 最短制动持续时间
            "speed_signal": "VehicleSpeed",            # 车速信号名
            "stop_speed_threshold": 1.0,               # 停车速度阈值 (km/h)
        }

    def detect(self, signals: Dict[str, pd.Series]) -> List[dict]:
        """
        检测制动事件

        Returns:
            制动事件列表，每个事件包含:
            - start_time: 制动开始时间
            - end_time: 制动结束时间
            - initial_speed: 初始车速
            - trigger_type: 触发方式 (brake_pedal / deceleration)
        """
        events = []

        # 方式1：基于刹车踏板信号
        brake_signal_name = self.config.get("brake_trigger_signal")
        if brake_signal_name and brake_signal_name in signals:
            events = self._detect_by_brake_pedal(signals[brake_signal_name], signals)
        else:
            # 方式2：基于减速度信号
            decel_signal_name = self.config.get("decel_trigger_signal")
            if decel_signal_name and decel_signal_name in signals:
                events = self._detect_by_deceleration(signals[decel_signal_name], signals)

        print(f"  检测到 {len(events)} 个制动事件")
        return events

    def _detect_by_brake_pedal(self, brake_signal: pd.Series,
                                signals: Dict[str, pd.Series]) -> List[dict]:
        """基于刹车踏板信号检测制动事件"""
        events = []
        threshold = self.config["brake_trigger_threshold"]
        speed_signal = signals.get(self.config["speed_signal"])

        # 找到刹车踏板踩下的区间
        is_braking = brake_signal > threshold
        # 找上升沿和下降沿
        rising = is_braking & ~is_braking.shift(1, fill_value=False)
        falling = ~is_braking & is_braking.shift(1, fill_value=False)

        start_times = brake_signal.index[rising]
        end_times = brake_signal.index[falling]

        for i, (start, end) in enumerate(zip(start_times, end_times)):
            duration = (end - start).total_seconds()
            if duration < self.config["min_duration_s"]:
                continue

            initial_speed = None
            if speed_signal is not None:
                # 取制动开始前0.5s的车速作为初始车速
                pre_start = start - pd.Timedelta(seconds=0.5)
                mask = (speed_signal.index >= pre_start) & (speed_signal.index <= start)
                if mask.any():
                    initial_speed = float(speed_signal[mask].iloc[0])

            events.append({
                "start_time": start,
                "end_time": end,
                "initial_speed": initial_speed,
                "trigger_type": "brake_pedal"
            })

        return events

    def _detect_by_deceleration(self, decel_signal: pd.Series,
                                 signals: Dict[str, pd.Series]) -> List[dict]:
        """基于减速度信号检测制动事件"""
        events = []
        threshold = self.config["decel_trigger_threshold"]
        speed_signal = signals.get(self.config["speed_signal"])

        is_braking = decel_signal < threshold
        rising = is_braking & ~is_braking.shift(1, fill_value=False)
        falling = ~is_braking & is_braking.shift(1, fill_value=False)

        start_times = decel_signal.index[rising]
        end_times = decel_signal.index[falling]

        for i, (start, end) in enumerate(zip(start_times, end_times)):
            duration = (end - start).total_seconds()
            if duration < self.config["min_duration_s"]:
                continue

            initial_speed = None
            if speed_signal is not None:
                pre_start = start - pd.Timedelta(seconds=0.5)
                mask = (speed_signal.index >= pre_start) & (speed_signal.index <= start)
                if mask.any():
                    initial_speed = float(speed_signal[mask].iloc[0])

            events.append({
                "start_time": start,
                "end_time": end,
                "initial_speed": initial_speed,
                "trigger_type": "deceleration"
            })

        return events


# ============================================================
# 指标计算模块
# ============================================================

class MetricsCalculator:
    """计算测试指标"""

    def __init__(self, config: dict = None):
        self.config = config or {
            "decel_signal": "VehicleDecel",        # 减速度信号
            "speed_signal": "VehicleSpeed",        # 车速信号
            "steering_signal": "SteeringAngle",    # 转向角信号
            "slip_signal": "WheelSlip",            # 车轮滑移率信号
            "mu_signal": "RoadMu",                 # 路面附着系数信号
            "slip_lock_threshold": 0.8,            # 滑移率抱死阈值
        }

    def calculate_abs_metrics(self, signals: Dict[str, pd.Series],
                               event: dict) -> dict:
        """
        计算单次ABS制动测试的指标

        Args:
            signals: 信号字典
            event: 制动事件（含start_time, end_time）

        Returns:
            指标字典
        """
        start = event["start_time"]
        end = event["end_time"]

        # 截取制动区间内的信号
        def slice_signal(sig):
            mask = (sig.index >= start) & (sig.index <= end)
            return sig[mask]

        metrics = {}

        # 1. 平均减速度 (m/s²)
        decel_sig = signals.get(self.config["decel_signal"])
        if decel_sig is not None:
            decel_slice = slice_signal(decel_sig)
            if len(decel_slice) > 0:
                metrics["平均减速度"] = round(float(np.mean(np.abs(decel_slice.values))), 2)

        # 2. 制动距离 (m)
        speed_sig = signals.get(self.config["speed_signal"])
        if speed_sig is not None:
            speed_slice = slice_signal(speed_sig)
            if len(speed_slice) > 1:
                # 对速度积分计算距离（速度单位km/h → m/s）
                dt = np.diff(speed_slice.index.astype(np.int64) / 1e9)  # 转为秒
                speed_ms = speed_slice.values[:-1] / 3.6  # km/h → m/s
                distance = float(np.sum(speed_ms * dt))
                metrics["制动距离"] = round(distance, 1)

        # 3. 减速度峰谷差值 (m/s²)
        if decel_sig is not None:
            decel_slice = slice_signal(decel_sig)
            if len(decel_slice) > 2:
                abs_decel = np.abs(decel_slice.values)
                # 计算相邻峰谷差
                peaks = []
                for i in range(1, len(abs_decel) - 1):
                    if abs_decel[i] > abs_decel[i-1] and abs_decel[i] > abs_decel[i+1]:
                        peaks.append(abs_decel[i])
                    elif abs_decel[i] < abs_decel[i-1] and abs_decel[i] < abs_decel[i+1]:
                        peaks.append(abs_decel[i])
                if len(peaks) >= 2:
                    peak_to_valley = [abs(peaks[i] - peaks[i+1]) for i in range(len(peaks)-1)]
                    metrics["减速度峰谷差值"] = round(float(max(peak_to_valley)), 1)

        # 4. 转向修正角 (deg)
        steering_sig = signals.get(self.config["steering_signal"])
        if steering_sig is not None:
            steering_slice = slice_signal(steering_sig)
            if len(steering_slice) > 0:
                # 转向修正角 = 方向盘最大偏转 - 最小偏转
                max_angle = float(np.max(np.abs(steering_slice.values)))
                metrics["转向修正角"] = round(max_angle, 1)

        # 5. 车轮抱死时间 (s)
        slip_sig = signals.get(self.config["slip_signal"])
        if slip_sig is not None:
            slip_slice = slice_signal(slip_sig)
            if len(slip_slice) > 0:
                threshold = self.config["slip_lock_threshold"]
                locked = np.abs(slip_slice.values) > threshold
                if locked.any():
                    dt = (slip_slice.index[-1] - slip_slice.index[0]).total_seconds() / len(slip_slice)
                    lock_time = float(np.sum(locked) * dt)
                    metrics["车轮抱死时间"] = round(lock_time, 2)
                else:
                    metrics["车轮抱死时间"] = 0.0

        # 6. 附着系数利用率 (%)
        mu_sig = signals.get(self.config["mu_signal"])
        if decel_sig is not None and mu_sig is not None:
            mu_slice = slice_signal(mu_sig)
            decel_slice = slice_signal(decel_sig)
            if len(mu_slice) > 0 and len(decel_slice) > 0:
                avg_mu = float(np.mean(mu_slice.values))
                avg_decel = float(np.mean(np.abs(decel_slice.values)))
                if avg_mu > 0:
                    utilization = (avg_decel / (avg_mu * 9.81)) * 100
                    metrics["附着系数利用率"] = round(min(utilization, 100), 0)

        # 7. 初始车速
        if event.get("initial_speed") is not None:
            metrics["初始车速"] = round(event["initial_speed"], 1)

        return metrics

    def calculate_averages(self, test_runs: List[dict]) -> dict:
        """计算多次测试的平均值"""
        if not test_runs:
            return {}

        numeric_keys = ["平均减速度", "制动距离", "减速度峰谷差值",
                        "转向修正角", "车轮抱死时间", "附着系数利用率"]

        average = {}
        for key in numeric_keys:
            values = [run.get(key) for run in test_runs if run.get(key) is not None]
            if values:
                average[key] = round(float(np.mean(values)), 2)

        # 主观评分取平均
        scores = [run.get("主观评分") for run in test_runs if run.get("主观评分") is not None]
        if scores:
            average["主观评分"] = round(float(np.mean(scores)), 1)

        # 结论：所有通过则平均通过
        conclusions = [run.get("结论") for run in test_runs if run.get("结论")]
        if all(c == "通过" for c in conclusions):
            average["结论"] = "通过"
        elif any(c == "通过" for c in conclusions):
            average["结论"] = "有条件通过"
        else:
            average["结论"] = "失败"

        return average


# ============================================================
# 事件分类模块
# ============================================================

class EventClassifier:
    """将制动事件按车速和测试类型分类"""

    # 车速分档（km/h）
    SPEED_BANDS = [
        {"label": "50",  "min": 45, "max": 55},
        {"label": "80",  "min": 75, "max": 85},
        {"label": "100", "min": 95, "max": 105},
    ]

    def __init__(self, config: dict = None):
        self.config = config or {}

    def classify_by_speed(self, events: List[dict]) -> Dict[str, List[dict]]:
        """按初始车速分档"""
        groups = {}
        for event in events:
            speed = event.get("initial_speed")
            if speed is None:
                continue
            for band in self.SPEED_BANDS:
                if band["min"] <= speed <= band["max"]:
                    label = band["label"]
                    if label not in groups:
                        groups[label] = []
                    groups[label].append(event)
                    break
        return groups

    def group_into_test_items(self, events_by_speed: Dict[str, List[dict]],
                               max_runs_per_speed: int = 3) -> List[dict]:
        """
        将事件组织成test_items格式

        每个车速取最多N次测试，组成一个test_item
        """
        test_items = []
        for speed_label, events in sorted(events_by_speed.items(), key=lambda x: int(x[0])):
            # 取前N次
            selected = events[:max_runs_per_speed]
            test_item = {
                "车速": int(speed_label),
                "test_runs": [],
                "average": {},
                "requirements": {}  # 从test_requirements.json填充
            }
            for i, event in enumerate(selected):
                run = {
                    "序号": i + 1,
                    # 指标由MetricsCalculator填充
                }
                test_item["test_runs"].append(run)

            test_items.append(test_item)

        return test_items


# ============================================================
# 主流程
# ============================================================

def process_mf4(mf4_path: str, signal_config_path: str,
                output_path: str, requirements_path: str = None,
                auto_detect: bool = True,
                event_config: dict = None) -> dict:
    """
    完整的MF4处理流程

    Args:
        mf4_path: MF4文件路径
        signal_config_path: 信号映射配置文件路径
        output_path: 输出JSON文件路径
        requirements_path: 测试要求配置文件路径
        auto_detect: 是否自动检测制动事件
        event_config: 事件检测配置

    Returns:
        处理后的数据字典
    """
    # 1. 加载配置
    with open(signal_config_path, 'r', encoding='utf-8') as f:
        signal_config = json.load(f)

    requirements = {}
    if requirements_path and os.path.exists(requirements_path):
        with open(requirements_path, 'r', encoding='utf-8') as f:
            requirements = json.load(f)

    # 2. 加载MF4文件并提取信号
    extractor = SignalExtractor(mf4_path, signal_config)
    extractor.load()

    # 列出可用信号（调试用）
    print("\n可用信号（含关键词过滤）：")
    for keyword in ["brake", "speed", "decel", "steering", "slip", "mu"]:
        matched = extractor.list_signals(keyword)
        if matched:
            print(f"  [{keyword}] 匹配 {len(matched)} 个信号:")
            for s in matched[:5]:
                print(f"    - {s}")
            if len(matched) > 5:
                print(f"    ... 还有 {len(matched) - 5} 个")

    # 提取映射的信号
    signals = extractor.extract_mapped_signals()
    print(f"\n成功提取 {len(signals)} 个信号")

    # 3. 检测制动事件
    detector = BrakeEventDetector(event_config)
    events = detector.detect(signals)

    # 4. 分类事件
    classifier = EventClassifier()
    events_by_speed = classifier.classify_by_speed(events)
    print(f"\n按车速分档:")
    for speed, evts in sorted(events_by_speed.items()):
        print(f"  {speed} km/h: {len(evts)} 个事件")

    # 5. 计算指标
    calculator = MetricsCalculator()

    # 构建输出数据
    output_data = {
        "project_name": "【项目名称 - 从MF4解析】",
        "report_id": f"RPT-{datetime.now().strftime('%Y%m%d')}-001",
        "report_date": datetime.now().strftime("%Y-%m-%d"),
        "source_file": os.path.basename(mf4_path),
        "parse_time": datetime.now().isoformat(),
    }

    # 处理ABS直线制动
    test_items = []
    for speed_label, speed_events in sorted(events_by_speed.items(), key=lambda x: int(x[0])):
        selected = speed_events[:3]  # 取前3次
        test_runs = []
        for i, event in enumerate(selected):
            metrics = calculator.calculate_abs_metrics(signals, event)
            run = {"序号": i + 1}
            run.update(metrics)
            run["主观评分"] = None  # 需人工填写
            run["结论"] = "【待判定】"
            test_runs.append(run)

        average = calculator.calculate_averages(test_runs)

        # 从requirements配置获取要求值
        req_key = "abs_straight_braking"
        req = requirements.get(req_key, {}).get("requirements", {})

        test_item = {
            "车速": int(speed_label),
            "test_runs": test_runs,
            "average": average,
            "requirements": req
        }
        test_items.append(test_item)

    if test_items:
        output_data["abs_straight_braking"] = {
            "test_conditions": {
                "测试路面": "【待确认】",
                "路面附着系数": None,
                "测试温度": None,
                "测试湿度": None
            },
            "test_items": test_items,
            "subjective_evaluation": "【待填写】"
        }

    # 6. 保存结果
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output_data, f, ensure_ascii=False, indent=2)
    print(f"\n✓ 解析结果已保存: {output_path}")

    extractor.close()

    return output_data


def main():
    parser = argparse.ArgumentParser(description='MF4数据解析器 - 提取CAN信号并计算测试指标')
    parser.add_argument('--input', '-i', required=True, help='MF4文件路径')
    parser.add_argument('--config', '-c', default='config/signal_mapping.json',
                        help='信号映射配置文件路径')
    parser.add_argument('--output', '-o', default='output/parsed_data.json',
                        help='输出JSON文件路径')
    parser.add_argument('--requirements', '-r', default='config/test_requirements.json',
                        help='测试要求配置文件路径')
    parser.add_argument('--auto-detect', action='store_true', default=True,
                        help='自动检测制动事件')
    parser.add_argument('--list-signals', action='store_true',
                        help='仅列出MF4文件中的可用信号')
    parser.add_argument('--filter', default=None,
                        help='信号过滤关键词')

    args = parser.parse_args()

    if args.list_signals:
        # 仅列出信号
        extractor = SignalExtractor(args.input, {})
        extractor.load()
        signals = extractor.list_signals(args.filter)
        print(f"\n共 {len(signals)} 个信号:")
        for s in signals:
            print(f"  {s}")
        extractor.close()
        return

    # 完整处理流程
    data = process_mf4(
        mf4_path=args.input,
        signal_config_path=args.config,
        output_path=args.output,
        requirements_path=args.requirements,
        auto_detect=args.auto_detect
    )

    # 打印摘要
    print("\n" + "=" * 60)
    print("解析摘要")
    print("=" * 60)
    abs_data = data.get("abs_straight_braking", {})
    test_items = abs_data.get("test_items", [])
    for item in test_items:
        speed = item.get("车速", "?")
        runs = item.get("test_runs", [])
        print(f"\n  {speed} km/h - {len(runs)} 次测试:")
        for run in runs:
            decel = run.get("平均减速度", "—")
            dist = run.get("制动距离", "—")
            print(f"    第{run.get('序号', '?')}次: 减速度={decel} m/s², 距离={dist} m")


if __name__ == '__main__':
    main()
