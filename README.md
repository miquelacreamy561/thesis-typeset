# 论文排版工具

面向通用论文场景的 Windows 桌面排版工具。默认内置 **华南农业大学 2024 本科毕业论文** 配置，也支持通过 YAML 配置适配其他学校。

这个版本更强调“直接好用”：
- 参数集中在一个桌面工作台里完成
- 常见论文格式问题可以一次处理
- 操作路径更清晰，适合赶论文时快速使用

![界面截图](image.png)

## 当前版本亮点

### 1. 一套流程处理常见论文排版任务
- 支持 `.docx`、`.doc`、`.txt`、`.md`、`.tex` 输入，统一输出 `.docx`
- 页面、正文、标题、页眉页码、目录、图表、封面声明都能集中配置
- 支持配置保存和加载，方便同校同学复用模板

### 2. 更适合桌面效率工具的界面
- 左侧导航、右侧配置区、底部操作区分工更清楚
- 长页面可以直接滚动，不用反复拖滚动条
- 主操作按钮和运行日志更容易找到
- 界面语言和交互都更偏实用，不做花哨装饰

### 3. 覆盖论文里最麻烦的一批格式问题
- 标题层级、编号与段距统一
- 页码、页眉、前置页与正文分区处理
- 图表题注、三线表、参考文献缩进等高频问题可配置
- 支持封面、声明页和特殊标题映射

## 主要优点

- 不只是改字体字号，而是把论文里常见的排版环节放到一套流程里处理
- 对“标题编号乱、目录不一致、页码不对、题注不规范、参考文献缩进错误”这类问题更有针对性
- 上手门槛低，适合先用默认配置直接跑，再按学校要求微调
- 既能点界面，也能走命令行，方便个人使用和分发给同学

## 使用时需要知道的边界

1. **仍然建议人工复核最终文档**
   - 这个工具能大幅减少重复劳动，但不能替代最终审稿和格式核对。

2. **不同学校要求差异大时，仍需要自己调配置**
   - 默认配置偏向 SCAU 2024。
   - 换学校时，封面、声明、摘要页、特殊标题等通常需要按本校要求调整。

3. **原稿越规范，自动处理效果越稳定**
   - 如果标题、题注、结构写法非常随意，工具可能无法完全按预期识别。

4. **部分能力依赖本机环境**
   - `.doc` 转换、目录更新等场景在 Windows + Word 环境下效果更完整。
   - `.txt/.md/.tex` 转换需要 Pandoc。

5. **复杂版面仍可能需要手工微调**
   - 比如图片内部文字、复杂表格局部布局、特殊文本框等，不属于完全自动处理范围。

## 适用场景

- 论文初稿已经写完，需要统一格式
- 学校有一套明确规范，但手工调整太耗时
- 同一学院或同学之间希望共享一套配置模板
- 想先快速生成一版规范文档，再做最后人工检查

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

当前依赖：
- `python-docx`
- `pyyaml`
- `pywin32`
- `ttkbootstrap`

### 2. 启动 GUI

```bash
python run_gui.py
```

### 3. 命令行使用

```bash
python thesis_format_cli.py --input 论文.docx
python thesis_format_cli.py --input 论文.docx --output 论文_规范版.docx
python thesis_format_cli.py --input 论文.docx --config thesis_config.yaml
python thesis_format_cli.py --dump-config
```

## 输入与依赖说明

| 输入格式 | 说明 | 额外依赖 |
|----------|------|----------|
| `.docx` | 直接进入格式化流程 | 无 |
| `.doc` | 先转 `.docx` | 需要 Microsoft Word |
| `.txt` | 预处理后经 Pandoc 转 `.docx` | 需要 Pandoc |
| `.md` | Pandoc 转 `.docx` | 需要 Pandoc |
| `.tex` | Pandoc 转 `.docx` | 需要 Pandoc |

## 项目结构

| 路径 | 说明 |
|------|------|
| `thesis_gui.py` | 桌面 GUI |
| `run_gui.py` | GUI 启动脚本 |
| `thesis_format_cli.py` | CLI / GUI 统一入口 |
| `thesis_runner.py` | 输入转换、格式化、后处理总流水线 |
| `thesis_config.py` | 默认配置与 YAML 加载 |
| `thesis_formatter/` | 核心排版逻辑 |
| `word_postprocess.py` | Word 后处理 |
| `defaults/scau_2024.yaml` | 默认学校配置 |
| `tests/` | 自动化测试 |

## 测试

```bash
python -m unittest discover -s tests
```

## 打包

项目内保留了 `build_exe.bat` 和 `thesis-format.spec`，可继续用于 Windows exe 打包。

## 许可证

GPL-3.0
