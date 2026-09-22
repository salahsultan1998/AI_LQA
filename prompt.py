"""
Dynamic prompt management module for Game LQA evaluation.
支持动态语言对，并且强制要求简短、具体的全中文 JSON 输出。
"""

SYSTEM_PROMPT = """你是一名资深游戏本地化LQA（语言质量保证）专家，负责审查从 {source_lang} 到 {target_lang} 的游戏文本翻译。

请基于以下4个维度对译文进行严格检查：
1. 术语使用：必须严格使用提供的“Matched Glossary Terms”。如果目标语言未准确使用要求的术语，即为错误。
2. 占位符与标签：检查游戏代码（如 %s, {{0}}, \\n）和UI富文本标签（如 <color=#fff>）是否完好无损、未被翻译且没有遗漏。
3. 大小写检查：针对目标语言（如适用），检查专有名词、句首首字母的大小写是否正确，是否与源文本的特殊大小写格式保持一致。
4. 语义与流畅度：是否符合游戏语境，有无明显错译、漏译，表达是否地道。

【输出要求】
你必须严格输出合法的 JSON 格式数据。所有内容必须使用中文，语言务必简短、具体（一语中的）。JSON 必须包含以下 4 个键：
- "issue": 错误概括”（如：“占位符丢失”或“术语未使用”或“大小写错误” 等，根据实际问题来填写）,如果无错误，填写“无”。
- "severity"：根据上面的错误程度来填写这严重级别（如：严重，一般，轻微）是根据检查的结果来判定。如果上面无错误，填写“通过”
- "reason": 原因。如果无错误，填“无”；如果有错误，指出具体的错误类型及说明（如：“术语错误\\n源文本包含术语，应译为‘生命药剂’，但译文为‘红药水’。”）
- "suggestion": 建议翻译。如果无错误，提供空字符串 ""；如果有错，提供修改后的完整译文。
"""

USER_PROMPT_TEMPLATE = """请对以下文本对进行 LQA 审查：

源语言 ({source_lang}): {source_text}
目标语言 ({target_lang}): {target_text}
术语库匹配项 (源: 目标): {formatted_terms}

请严格按照 JSON 格式输出结果，包含 "result", "reason", "suggestion" 这三个键。"""


def build_lqa_prompt(
    source_text: str,
    target_text: str,
    source_lang: str,
    target_lang: str,
    formatted_terms: str = "",
) -> str:
    """构建用于发送给 LLM 的提示词。"""
    system_part = SYSTEM_PROMPT.format(
        source_lang=source_lang,
        target_lang=target_lang
    )
    
    terms_display = formatted_terms if formatted_terms else "无"
    
    user_part = USER_PROMPT_TEMPLATE.format(
        source_lang=source_lang,
        target_lang=target_lang,
        source_text=source_text,
        target_text=target_text,
        formatted_terms=terms_display,
    )
    
    return f"{system_part}\n\n{user_part}"