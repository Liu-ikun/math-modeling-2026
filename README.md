# 2026 CUMCM 微电网储能优化（Problem 1–4）

全国大学生数学建模竞赛 2026 年赛题的求解代码与结果。题目围绕**光伏 + 储能微电网**的购电与充放电策略：
在分时电价、光伏出力、负荷数据下，用线性规划（LP）与模型预测控制（MPC）最小化购电成本，
并给出无储能基线、完美信息下界、成本阶梯分解与灵敏度分析作为侧面验证。

## 目录结构

```
2026_CUMCM/
├── Problem_1/              问题1：典型日单日最优购电（144 时段 LP）+ 无储能基线
├── Problem_2/              问题2：全年 365 天全局 LP（完美信息）+ 逐日滚动对照
├── Problem_3/              问题3：含预报误差的 MPC 滚动优化 + 成本阶梯 + 更新频率实验
├── Problem_4/              问题4：附件4 变电价 α_d 下的全局 LP（Q4-2）与 MPC（Q4-3）
├── Sensitivity_Analysis/   灵敏度分析
├── Shared_Utils/           数据清洗与公共工具（preprocessing.py 生成所有 processed 数据）
└── Docs/                   论文、文献与过程笔记
```

每个 `Problem_X` 内部统一为：

```
Problem_X/
├── data/raw/               原始附件（附件1~4）
├── data/processed/         清洗后数据（由 preprocessing.py 生成，不入库）
├── src/                    main.py 主流程 + model.py 模型 + algorithm.py 算法
└── outputs/
    ├── tables/             结果表（xlsx / csv）
    ├── figures/            论文插图（png）
    └── report/summary.md   求解摘要（含三组侧面验证指标）
```

## 环境准备

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

主要依赖：`pandas`、`numpy`、`PuLP`（LP 求解）、`matplotlib`、`scipy`、`scikit-learn`、`lifelines`。

## 数据准备

仓库**不包含** `data/processed/`（中间数据体积大且可完全复现）。克隆后先跑一次清洗：

```powershell
cd 2026_CUMCM/Shared_Utils
python preprocessing.py
```

该脚本读取各 `Problem_X/data/raw/` 下的附件1~4，统一时间轴、宽表转长表、附件3 做 10 min 线性插值，
并按问题合并后写出 `data/processed/`，同时生成 `cleaning_report.md` 与 `preprocessing.log`。

## 运行

每个问题相互独立，直接运行各自的 `main.py` 即可（路径均相对脚本自身解析，可从任意目录启动）：

```powershell
cd 2026_CUMCM/Problem_1/src
python main.py
```

结果写入同级 `outputs/`：`tables/` 结果表、`figures/` 插图、`report/summary.md` 摘要。

## 主要结果（节选）

| 问题 | 结论 |
| --- | --- |
| 问题1（典型日） | 购电费 35126.95 元，较无储能 48052.05 元节省 26.9% |
| 问题2（全年全局 LP） | 13751266.38 元，较无储能节省 4547325.98 元（24.85%），弃光消纳率提升 68.27% |
| 问题3（MPC） | 提交策略 19674667.34 元，滚动更新信息价值 181681.78 元 |
| 问题4（变电价） | 全局最优 14494771.23 元；MPC 提交 19905270.72 元，时移修复价值 607888.30 元 |

## 版本管理说明

`.gitignore` 已排除：`.venv/`、`__pycache__/`、`.idea/`、`**/data/processed/` 及系统临时文件。
原始数据（`data/raw/`）、结果表、插图和摘要均纳入版本管理。
