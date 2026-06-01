# 讲义 Schema

## 期望格式

讲义是 docx 或 PDF。结构以 **Heading 1 + Normal 段落** 组织。

## 解析后的 JSON Schema

```typescript
interface Handout {
  title: string;            // "第一讲 分解质因数"
  lecture_num: string;      // "第1讲" / "第13讲"
  
  // 课程目标 + 知识阶梯（可能合并在课程目标 Heading 下）
  objectives: {
    知识技能: string;
    数学能力: string;
    思想方法: string;
    情感态度: string;
  };
  knowledge_ladder: string | null;
  
  // 模块（讲义里的物理模块，可能跟教学设计稿的"教学模块"不一致）
  modules: HandoutModule[];
  
  // 拓展（创新挑战）
  innovation: HandoutItem | null;
}

interface HandoutModule {
  module_num: number;
  title: string;            // 如 "分解质因数的方法"
  items: HandoutItem[];
}

interface HandoutItem {
  kind: "例题" | "练习";
  num: number;              // 1, 2, 3, ...
  tag: string | null;       // 如 "BYGF" / "2022 • 专项"
  body: string;             // 题干（含分项 (1)(2)(3)）
  solution: string | null;  // 题解（如果讲义有，通常没有，要从教学设计稿取）
}
```

## 解析规则

### Heading 识别

| Heading 级别 | 角色 |
|---|---|
| 首段（Normal）含"第N讲"或"第N节" | title + lecture_num |
| Heading 1 = "知识阶梯" | knowledge_ladder 起点 |
| Heading 1 = "课程目标" | objectives 起点 |
| Heading 1 = 其他 | module 起点（module title） |
| Heading 2 = "创新挑战" | innovation 起点 |

### Item 识别

在 module 内，每个 Normal 段落开头匹配以下正则识别 item：
- `^例题\s*(\d+)\s*(?:【(.+?)】)?(.+)` → kind=例题, num=1, tag="..."
- `^练习\s*(\d+)\s*(?:【(.+?)】)?(.+)` → kind=练习, num=1
- `^例\s*(\d+)` → 备用模式

题干持续到下一个 item 起始或下一个 Heading。

### 题号唯一性

- 例题编号必须从 1 开始连续递增
- 练习编号通常和例题编号配对（练1 配 例1）
- 如果讲义里题号跳号，解析时报 warning

## 已验证的讲义样本

- `第1讲分解质因数.docx`：3 个模块 + 6 例题 + 6 练习 + 1 创新挑战
- `牛吃草问题讲义.pdf`：3 个模块 + 5 例题 + 5 练习
