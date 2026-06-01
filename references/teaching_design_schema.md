# 教学设计稿 Schema

## 期望格式

教学设计稿是 docx（推荐）或 markdown。结构以 **Heading 1/2/3 + 段落** 组织。

## 解析后的 JSON Schema

```typescript
interface TeachingDesign {
  title: string;              // 来自 Heading 1，如 "《分解质因数》110分钟教学设计"
  total_duration_min: number; // 从标题或元信息提取
  
  story_line: {
    title: string;            // Heading 2 "故事线：XXX" 后的文字
    background: string;       // 课程背景设定段落
    student_role: string;     // 学生扮演的角色（如"密码破解者"）
  } | null;
  
  segments: Segment[];
  
  meta: {
    time_allocation: string;
    story_recap: string[];
    key_points: string[];
    extension_challenges: string;
    teaching_highlights: string[];
  };
}

interface Segment {
  segment_num: string;        // "一" / "二" / "三"
  segment_type: "导入" | "模块" | "总结";
  module_title: string;       // 如 "分解质因数的方法"
  duration_min: number;       // 25
  subtitle: string | null;    // 如 "工具修炼：短除法秘籍"
  
  activities: Activity[];
}

interface Activity {
  activity_type: "情境创设" | "探究活动" | "技能学习" 
               | "例题解析" | "学生练习" | "故事推进" 
               | "核心概念讲解" | "方法总结" | "方法迁移"
               | "题目分析" | "破解过程" | "其他";
  duration_min: number | null;
  raw_text: string;           // markdown 原文，保留缩进和列表
  bullet_items: string[];     // 解析出的 numbered/bulleted items
  
  // 如果引用讲义题目
  references: HandoutItemRef[];
}

interface HandoutItemRef {
  kind: "例题" | "练习";
  num: number;                // 题号
}
```

## 解析规则

### Heading 识别

| Heading 级别 | 角色 |
|---|---|
| Heading 1 | 教学设计标题（通常只 1 个） |
| Heading 2 | segment（导入/模块/总结） |
| Heading 3 | meta 或 segment 内部小节 |

### Segment 识别

`Heading 2` 文字含以下任一关键词时识别为 segment：
- "导入" → segment_type = "导入"
- "模块" → segment_type = "模块"
- "总结" / "拓展" → segment_type = "总结"
- "亮点" / "教学准备" / "预期效果" / "故事线" → meta（不算 segment）

从 segment 标题中提取 `duration_min`（如"25分钟"→ 25）。

### Activity 识别

每个 segment 内，Normal 段落按以下规则切分成 activities：

- 含关键词的段落作为 activity 起始：
  - "情境创设：" / "探究活动" / "技能学习" / "例题解析" / "学生练习"
  - "故事推进" / "核心概念讲解" / "方法总结" / "方法迁移"
- activity 持续到下一个 activity 起始或 segment 结束
- 末尾的"（X分钟）"提取为 duration_min

### references 提取

在 activity 的 raw_text 里 grep `例题(\d+)` 或 `例(\d+)` 或 `练习(\d+)` 或 `练(\d+)`，提取数字作为 num。

## 已验证的教学设计稿样本

- `分解质因数110分钟教学设计.docx`：6 个 Heading 2（故事线/导入/模块一/二/三/四/总结），结构规范，可正确解析。
