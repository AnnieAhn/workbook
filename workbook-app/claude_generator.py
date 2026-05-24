"""
Claude API를 이용해 영어 지문을 분석하고 워크북 전체 내용을 생성합니다.
"""
import json
import anthropic

SYSTEM_PROMPT = """당신은 한국 고등학교 수능/내신 영어 전문 교사입니다.
영어 지문을 분석하여 학생이 1시간 이상 집중 학습할 수 있는 완전정복 워크북 내용을 생성합니다.
반드시 유효한 JSON만 반환하세요. 마크다운 코드블록 없이 순수 JSON만 출력하세요."""

WORKBOOK_PROMPT = """다음 영어 지문을 분석하고 워크북에 필요한 모든 내용을 JSON으로 생성하세요.

[영어 지문]
{passage}

[요구사항]
- 핵심 어휘: {vocab_count}개 (지문 등장 순서)
- T/F 문제: {tf_count}개
- 어법 판단: {"포함" if include_grammar else "제외"}
- 서술형: {"포함" if include_essay else "제외"}

아래 JSON 구조를 정확히 따르세요:

{{
  "analysis": {{
    "topic_kor": "한국어 주제 (한 문장)",
    "main_idea_kor": "한국어 요지 (필자가 전달하려는 핵심 메시지)",
    "passage_type": "논설문 또는 설명문 또는 실용문 또는 서사문",
    "title_candidates": ["영어 제목 후보1", "영어 제목 후보2", "영어 제목 후보3"],
    "structure_type": "두괄식 또는 미괄식 또는 대조 또는 열거 또는 인과",
    "paragraph_summaries": ["문단1 핵심내용 (한국어)", "문단2 핵심내용", "..."],
    "key_connectors": [
      {{"expression": "연결어/핵심표현", "location": "위치 설명", "function": "기능 설명"}}
    ]
  }},
  "vocabulary": [
    {{
      "no": 1,
      "word": "단어/표현",
      "pos": "품사 (v./n./adj./adv./phr. 등)",
      "meaning": "한국어 뜻",
      "example": "지문에서 실제 사용된 예문 (가능한 짧게)",
      "synonym": "동의어1, 동의어2"
    }}
  ],
  "vocab_test_words": [
    {{"meaning_kor": "한국어 뜻", "answer": "영어단어"}}
  ],
  "tf_questions": [
    {{
      "no": 1,
      "statement": "T/F 판단할 영어 문장",
      "answer": "T",
      "evidence": "지문에서 근거가 되는 문장 (원문 그대로)"
    }}
  ],
  "content_choice_q1": {{
    "question": "According to the passage, which of the following is TRUE?",
    "choices": ["① 선택지1", "② 선택지2", "③ 선택지3 (정답)", "④ 선택지4", "⑤ 선택지5"],
    "answer": "③",
    "explanations": {{
      "1": "① 오답 이유",
      "2": "② 오답 이유",
      "3": "③ 정답 근거 (지문 문장 인용)",
      "4": "④ 오답 이유",
      "5": "⑤ 오답 이유"
    }}
  }},
  "content_choice_q2": {{
    "question": "What can be inferred from the passage?",
    "choices": ["① 선택지1", "② 선택지2", "③ 선택지3", "④ 선택지4 (정답)", "⑤ 선택지5"],
    "answer": "④",
    "explanations": {{
      "1": "① 오답 이유",
      "2": "② 오답 이유",
      "3": "③ 오답 이유",
      "4": "④ 정답 근거",
      "5": "⑤ 오답 이유"
    }}
  }},
  "subject_type": "주제 또는 요지 또는 제목",
  "subject_q1": {{
    "question": "다음 글의 [주제/요지/제목]으로 가장 적절한 것은?",
    "choices": ["① 선택지1", "② 선택지2", "③ 선택지3", "④ 선택지4", "⑤ 선택지5 (정답)"],
    "answer": "⑤",
    "difficulty": "표준",
    "explanations": {{
      "1": "① 오답 이유",
      "2": "② 오답 이유",
      "3": "③ 오답 이유",
      "4": "④ 오답 이유",
      "5": "⑤ 정답 근거"
    }}
  }},
  "subject_q2": {{
    "question": "다음 글의 [주제/요지/제목]으로 가장 적절한 것은? (선택지를 모두 신중하게 검토하세요)",
    "choices": ["① 선택지1", "② 선택지2 (정답)", "③ 선택지3", "④ 선택지4", "⑤ 선택지5"],
    "answer": "②",
    "difficulty": "고난도",
    "explanations": {{
      "1": "① 오답 이유",
      "2": "② 정답 근거",
      "3": "③ 오답 이유",
      "4": "④ 오답 이유",
      "5": "⑤ 오답 이유"
    }}
  }},
  "blank_q1": {{
    "label": "쉬움",
    "hint_kor": "힌트: 지문의 핵심 개념어",
    "passage_with_blank": "빈칸 포함 문장 (______으로 표시)",
    "choices": ["① 선택지1", "② 선택지2 (정답)", "③ 선택지3", "④ 선택지4", "⑤ 선택지5"],
    "answer": "②",
    "explanation": "정답 근거 설명"
  }},
  "blank_q2": {{
    "label": "보통",
    "hint_kor": "힌트: 전후 문맥 파악 필요",
    "passage_with_blank": "빈칸 포함 문장",
    "choices": ["① 선택지1", "② 선택지2", "③ 선택지3 (정답)", "④ 선택지4", "⑤ 선택지5"],
    "answer": "③",
    "explanation": "정답 근거 설명"
  }},
  "blank_q3": {{
    "label": "어려움",
    "hint_kor": "힌트: 지문 전체 논리 파악 필요",
    "passage_with_blank": "빈칸 포함 문장",
    "choices": ["① 선택지1", "② 선택지2", "③ 선택지3", "④ 선택지4 (정답)", "⑤ 선택지5"],
    "answer": "④",
    "explanation": "정답 근거 설명"
  }},
  "grammar_items": [
    {{
      "no": 1,
      "point": "문법 포인트 (예: 관계사절, 수동태 등)",
      "sentence": "지문에서 가져온 완전한 문장",
      "underline": "밑줄 칠 부분 (sentence에 포함된 텍스트)",
      "correct": true,
      "correction": "틀린 경우 올바른 형태, 맞으면 —",
      "explanation": "문법 해설 (한국어)"
    }}
  ],
  "short_answer_qs": [
    {{
      "no": 1,
      "question": "영어 질문문?",
      "hint_kor": "힌트 (한국어)",
      "model_answer": "모범 답안 (영어)"
    }}
  ],
  "summary_template": "이 글은 ①[    ]에 대해 설명한다. ②[    ]의 중요성을 강조하며, ③[    ]가 ④[    ]에 미치는 영향을 보여준다.",
  "summary_answers": ["①의 답", "②의 답", "③의 답", "④의 답"],
  "writing_tasks": [
    {{
      "no": 1,
      "korean": "영작할 한국어 문장",
      "hint_words": "( word1 / word2 / word3 )",
      "model_answer": "영어 모범 답안"
    }}
  ],
  "one_line_summary": {{
    "answer_part1": "The passage mainly discusses ...",
    "answer_part2": "because/by ..."
  }},
  "classcard": "단어1\\t한국어뜻1\\n단어2\\t한국어뜻2\\n..."
}}

중요 지침:
1. 모든 선택지는 지문과 관련 있어야 하며 그럴듯한 함정을 포함해야 합니다
2. T/F 정답은 T와 F가 골고루 섞여야 합니다 (F가 너무 적으면 안 됨)
3. 어법 문제 5개 중 1-2개는 실제로 틀린 문장으로 만드세요
4. 클래스카드 형식: 단어\\t뜻 (탭으로 구분, 줄바꿈으로 분리)
5. 반드시 유효한 JSON만 반환하세요"""


def generate_workbook_content(
    passage: str,
    api_key: str,
    vocab_count: int = 20,
    tf_count: int = 10,
    include_grammar: bool = True,
    include_essay: bool = True,
) -> dict:
    """
    Claude API를 호출해 워크북 전체 내용을 JSON으로 반환합니다.
    """
    client = anthropic.Anthropic(api_key=api_key)

    prompt = WORKBOOK_PROMPT.format(
        passage=passage,
        vocab_count=vocab_count,
        tf_count=tf_count,
        include_grammar=include_grammar,
        include_essay=include_essay,
    )

    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=8000,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": prompt}],
    )

    raw = message.content[0].text.strip()

    # JSON 블록이 있으면 제거
    if raw.startswith("```"):
        lines = raw.split("\n")
        raw = "\n".join(lines[1:-1] if lines[-1] == "```" else lines[1:])

    return json.loads(raw)
