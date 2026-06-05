# content.json / ai_scaffold.json Schema

## 通用字段（extract_handout.py 产出）
```jsonc
{
  "lecture_no": "第09讲", "title": "情节模板\n克服心魔",
  "lecture_type": "continuation | reading",   // 自动识别：含「固化情节」→continuation；含「阅读X篇」→reading
  "grade": "高中英语", "teacher": "主讲老师：",
  "preview": {"header": ["Parts","Points","Levels","Objectives"], "rows": [...]},   // 讲义 Preview 表
  "leading_in": "…",
  "parts": [{"part_label":"PART 1","title":"…","subtitle":"…"}, ...],
  "end": {"big":"本节课结束","small":"下节课我们\n再见啦～"}
}
```

## continuation（读后续写·主观题）
```jsonc
{
  "part1_plots": [   // 固化情节N：基础/升级 中英翻译（讲义【答案】中文/英文）
    {"index":1,"name":"坚持克服",
     "basic":{"zh":"…","en":"…","tag":"【知识标签】…"},
     "upgrade":{"zh":"…","en":"…","tag":"…"}}
  ],
  "part2": {        // 篇章实战（part3=篇章复用，结构相同）
    "source": "【2024.01·浙江卷】",
    "passage_en": ["① …","② …", ...],
    "prompts": {"para1":"Para 1: …","para2":"Para 2: …"},   // 续写提示句
    "details": [    // 三、情节描写（首段1-4/次段5-8）
      {"index":1,"group":"首段","label":"认可点赞","zh":"…","en":"…","tag":"…"}
    ],
    "passage_zh": ["…"]   // C类·AI翻译（可选，由 AI 步骤补入）
  }
}
```

## reading（阅读理解·客观题）
```jsonc
{
  "part1_zhuzhi":   {"title":"主旨辅助","source":"【…】","passage":[…],
                     "method_table":[["原文首尾段关键信息",""],["文章主旨",""]],
                     "questions":[Q...]},
  "part1_xuanxiang":{"title":"选项辅助","table":[[…4列…],…]},
  "part2": [ {"name":"阅读B篇","level":"⭐⭐⭐⭐","source":"【…】",
              "passage":[…], "questions":[Q...]} ],
  // Q = {"stem":"(1)…","options":["A．…","B．…",…],"answer":"D","analysis":"…","tag":"…"}
}
```

## ai_scaffold.json（C类·AI生成，须配《AI生成内容审核清单》）
```jsonc
{
  "plot_overview": {"title":"克服心魔类·固化情节","steps":["坚持克服",…]},
  "step_framework": ["步骤一·原文概括","步骤二·细节推理","步骤三·段首分析"],
  "guide_prompt": "你会怎么表达?",
  "part2": {            // part3 同
    "overall_summary":  "原文概括（AI）",
    "overall_direction":"后续走向/整体方向（AI）",
    "reasoning_q": {"context":"原文细节…","stem":"以下哪个后续情节推测更符合原文细节？",
                    "options":["正确项(=后续走向)","干扰项1","干扰项2"],"answer":0},
    "desc_types": ["动作描写","动作+环境","心理描写","情绪描写","环境描写","通用型", …]  // 按 details index 序
  }
}
```

## 解析容错（坑）
- 题目两种排布：题后紧跟答案块（主旨辅助式）/ 所有题后集中答案块（B/C/D篇式）——集中块首条【答案】为空(主题概述)须跳过再按序配对
- 源行可能混合：`无1⭐⭐⭐⭐【2025·…】passage第一句…` → 拆 星级/【题源】/语篇首段
- 知识标签=题库元数据 → 渲染进 PPT 备注，不上正文
