# pptx2md

将 PowerPoint (.pptx) 转换为 **LLM 友好的 Markdown** 格式。

专为大模型阅读优化：自动提取标题、正文、列表、表格，过滤图片/形状噪声，短标签聚合去重。

## 安装

```bash
pip install python-pptx
```

## 用法

```bash
# 输出到终端
python pptx2md_llm.py input.pptx

# 写入文件
python pptx2md_llm.py input.pptx -o output.md

# 同时导出图片
python pptx2md_llm.py input.pptx -o output.md -i imgs/

# 附加位置坐标信息
python pptx2md_llm.py input.pptx --with-position
```

## 输出示例

```markdown
# 演示文稿名称

幻灯片尺寸：33.9 × 19.1 cm | 共 9 页

---
## 第1页：网络安全新考验

标签：远程办公易被攻击 | 分支安全牵连总部 | 运营成本增加

---
## 第2页：解决方案概览

AI融合Secure SD-WAN解决方案：LAN/WAN融合部署

* 管理通道
* 用于控制器编排和下发网络配置
* 控制通道
* 用于分发站点之间的路由和隧道信息

标签：总部 | USG 网关 | MPLS | Internet | 分支A | 分支B

> **备注:** 这里是演讲者备注内容
```

## 设计思路

| 元素 | 处理方式 |
|---|---|
| 标题占位符 | → `## 第X页：标题` |
| 长文本框 (>15字) | 原样输出，保留 `*` 分点和缩进 |
| 短标签 (≤15字) | 去重后聚合为 `标签：A \| B \| C` |
| 表格 | → Markdown 表格 |
| 图片/形状/连接线 | 跳过（减少噪声） |
| 演讲者备注 | → `> **备注:** ...` |

## Python API

```python
from pptx2md_llm import pptx_to_markdown

md = pptx_to_markdown("input.pptx")
print(md)
```

## License

MIT
