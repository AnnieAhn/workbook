"""
📘 수능/내신 영어 완전정복 워크북 생성기
Streamlit 웹 앱

실행 방법:
  pip install -r requirements.txt
  streamlit run app.py

Streamlit Cloud 배포 시:
  - Secrets 에 ANTHROPIC_API_KEY 추가
"""
import os
import datetime
import streamlit as st
from claude_generator import generate_workbook_content
from workbook_builder import build_workbook_docx

# ── 페이지 설정 ────────────────────────────────────────
st.set_page_config(
    page_title="완전정복 워크북 생성기",
    page_icon="📘",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS ───────────────────────────────────────────────
st.markdown("""
<style>
    .main-title {
        font-size: 2rem;
        font-weight: 800;
        color: #1F4E79;
        margin-bottom: 0.2rem;
    }
    .sub-caption {
        color: #555;
        font-size: 0.95rem;
        margin-bottom: 1.5rem;
    }
    .info-box {
        background: #EBF3FB;
        border-left: 4px solid #2E74B5;
        padding: 0.8rem 1rem;
        border-radius: 4px;
        margin: 0.5rem 0;
        font-size: 0.9rem;
    }
    .section-card {
        background: #f8f9fa;
        border: 1px solid #dee2e6;
        border-radius: 8px;
        padding: 1rem;
        margin: 0.5rem 0;
    }
    .stDownloadButton > button {
        background-color: #1F4E79 !important;
        color: white !important;
        font-size: 1.1rem !important;
        padding: 0.7rem 2rem !important;
    }
</style>
""", unsafe_allow_html=True)

# ── 사이드바 ─────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ 설정")

    # API 키 (Streamlit Secrets → 환경변수 → 입력창 순)
    env_key = st.secrets.get("ANTHROPIC_API_KEY", "") or os.environ.get("ANTHROPIC_API_KEY", "")
    if env_key:
        api_key = env_key
        st.success("✅ API 키: 서버에서 자동 로드됨")
    else:
        api_key = st.text_input(
            "Anthropic API 키",
            type="password",
            placeholder="sk-ant-...",
            help="https://console.anthropic.com 에서 발급 | 무료 $5 크레딧 제공",
        )
        if not api_key:
            st.warning("API 키를 입력해야 워크북을 생성할 수 있습니다.")

    st.divider()
    st.markdown("### 📝 워크북 옵션")

    vocab_count = st.slider(
        "핵심 어휘 수",
        min_value=15, max_value=25, value=20, step=1,
        help="지문에서 추출할 핵심 어휘 개수"
    )
    tf_count = st.slider(
        "T/F 문제 수",
        min_value=7, max_value=12, value=10, step=1,
    )
    include_grammar = st.checkbox("어법 판단 문제 포함", value=True)
    include_essay   = st.checkbox("서술형 문제 포함", value=True)
    include_classcard = st.checkbox("📲 클래스카드 단어 세트 출력", value=True)

    st.divider()
    st.markdown("""
    **📌 포함 섹션**
    1. 지문 읽기
    2. 핵심 어휘 + 확인 테스트
    3. 구조 분석 (흐름도식)
    4. T/F + 내용일치
    5. 주제/요지/제목 + 빈칸 + 어법
    6. 서술형 (단답형·요약·영작)
    7. 심화 (나만의 문제·오답노트)
    8. 정답 및 해설 (오답 함정 해설 포함)
    """)


# ── 메인 화면 ─────────────────────────────────────────
st.markdown('<p class="main-title">📘 수능/내신 영어 완전정복 워크북 생성기</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-caption">영어 지문 하나를 붙여넣으면 A4 5~8페이지 분량의 워크북 Word 파일을 자동 생성합니다</p>', unsafe_allow_html=True)

st.markdown("""
<div class="info-box">
💡 <b>사용법</b>: ① 지문 정보 입력 → ② 지문 붙여넣기 → ③ 워크북 생성 버튼 클릭 → ④ .docx 다운로드
</div>
""", unsafe_allow_html=True)

st.divider()

# ── 지문 정보 ─────────────────────────────────────────
col1, col2, col3 = st.columns([3, 3, 1])
with col1:
    passage_no = st.text_input(
        "지문 번호 / 출처",
        placeholder="예: 올림포스 Chapter2 QN29 3번",
    )
with col2:
    topic_hint = st.text_input(
        "주제 힌트 (선택 — 파일명에 사용)",
        placeholder="예: 신체움직임과창의성",
    )
with col3:
    difficulty = st.selectbox("예상 난이도", ["중", "상", "최상"])

# ── 지문 입력 ─────────────────────────────────────────
passage = st.text_area(
    "📄 영어 지문을 붙여넣으세요",
    height=220,
    placeholder=(
        "By moving our bodies, we activate a deeply ingrained and mostly unconscious "
        "metaphor connecting dynamic motion with dynamic thinking...\n\n"
        "여기에 영어 지문 전체를 붙여넣으세요."
    ),
)

word_count = len(passage.split()) if passage else 0
if passage:
    st.caption(f"📊 단어 수: {word_count}개 | 예상 수능 지문 길이: 150~280 단어")

st.divider()

# ── 생성 버튼 ─────────────────────────────────────────
can_generate = bool(api_key and passage and passage.strip())
btn_label = "🚀 워크북 생성하기" if can_generate else "🚀 워크북 생성하기 (API 키 + 지문 필요)"

if st.button(btn_label, type="primary", use_container_width=True, disabled=not can_generate):

    if word_count < 50:
        st.error("지문이 너무 짧습니다. 수능/내신 영어 지문을 전체 붙여넣어 주세요.")
        st.stop()

    progress_bar = st.progress(0, text="Claude가 지문을 분석 중입니다...")

    try:
        # Step 1: Claude API 호출
        progress_bar.progress(15, text="🔍 지문 분석 및 어휘 추출 중...")
        with st.spinner(""):
            content = generate_workbook_content(
                passage=passage,
                api_key=api_key,
                vocab_count=vocab_count,
                tf_count=tf_count,
                include_grammar=include_grammar,
                include_essay=include_essay,
            )

        progress_bar.progress(65, text="📝 워크북 레이아웃 구성 중...")

        # Step 2: Word 파일 생성
        docx_bytes = build_workbook_docx(
            content=content,
            passage=passage,
            passage_no=passage_no,
            topic_hint=topic_hint or content.get("analysis", {}).get("topic_kor", "영어지문"),
            include_grammar=include_grammar,
            include_essay=include_essay,
        )

        progress_bar.progress(100, text="✅ 완료!")

        st.success("✅ 워크북 생성 완료!")

        # ── 다운로드 버튼 ─────────────────────────────
        today_str   = datetime.date.today().strftime("%Y-%m-%d")
        safe_topic  = (topic_hint or "워크북").replace(" ", "_")
        filename    = f"{today_str}_{safe_topic}_완전정복워크북_v1.docx"

        st.download_button(
            label="📥  워크북 다운로드 (.docx)",
            data=docx_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

        # ── 분석 요약 미리보기 ────────────────────────
        analysis = content.get("analysis", {})
        col_a, col_b = st.columns(2)

        with col_a:
            with st.expander("📊 지문 분석 결과", expanded=True):
                st.markdown(f"**주제**: {analysis.get('topic_kor','—')}")
                st.markdown(f"**요지**: {analysis.get('main_idea_kor','—')}")
                st.markdown(f"**유형**: {analysis.get('passage_type','—')}  |  **구조**: {analysis.get('structure_type','—')}")
                st.markdown(f"**출제 유형**: {content.get('subject_type','—')}")
                st.markdown("**핵심 어휘 (상위 5개)**")
                for v in content.get("vocabulary", [])[:5]:
                    st.markdown(f"- **{v.get('word','')}** ({v.get('pos','')}): {v.get('meaning','')}")

        with col_b:
            with st.expander("📋 생성된 문제 수", expanded=True):
                tf_cnt    = len(content.get("tf_questions", []))
                vocab_cnt = len(content.get("vocabulary", []))
                blank_cnt = sum(1 for k in ["blank_q1","blank_q2","blank_q3"] if content.get(k))
                gram_cnt  = len(content.get("grammar_items", [])) if include_grammar else 0
                essay_cnt = len(content.get("short_answer_qs", [])) if include_essay else 0
                st.markdown(f"""
| 섹션 | 문항 수 |
|------|--------|
| 핵심 어휘 | {vocab_cnt}개 |
| T/F 문제 | {tf_cnt}개 |
| 5지선다 (내용일치) | 2문제 |
| 5지선다 ({content.get('subject_type','유형')}) | 2문제 |
| 빈칸 추론 | {blank_cnt}문제 |
| 어법 판단 | {gram_cnt}개 |
| 서술형 | {essay_cnt}문제 |
""")

        # ── 클래스카드 출력 ──────────────────────────
        if include_classcard:
            classcard_text = content.get("classcard", "")
            if classcard_text:
                with st.expander("📲 클래스카드 단어 세트 (복사 후 클래스카드에 붙여넣기)", expanded=False):
                    st.markdown("""
<div class="info-box">
클래스카드 → 학습세트 만들기 → <b>일괄 입력</b> 탭 → 아래 내용 전체 복사 후 붙여넣기
</div>
""", unsafe_allow_html=True)
                    st.code(classcard_text, language=None)
                    st.download_button(
                        "📥 클래스카드 텍스트 다운로드",
                        data=classcard_text,
                        file_name=f"{today_str}_{safe_topic}_클래스카드.txt",
                        mime="text/plain",
                    )

    except Exception as e:
        progress_bar.empty()
        err_msg = str(e)
        if "api_key" in err_msg.lower() or "authentication" in err_msg.lower():
            st.error("❌ API 키가 잘못되었습니다. Anthropic Console에서 발급한 키를 확인해주세요.")
        elif "json" in err_msg.lower():
            st.error("❌ AI 응답 파싱 오류입니다. 다시 시도해주세요.")
            st.caption(f"상세 오류: {err_msg}")
        else:
            st.error(f"❌ 오류가 발생했습니다: {err_msg}")

# ── 하단 안내 ─────────────────────────────────────────
st.divider()
with st.expander("ℹ️ API 키 발급 방법 & 로컬 실행 방법"):
    st.markdown("""
### Anthropic API 키 발급 (무료 $5 크레딧)
1. [console.anthropic.com](https://console.anthropic.com) 접속
2. 회원가입 / 로그인
3. **API Keys** → **Create Key**
4. 생성된 키(`sk-ant-...`)를 위 입력창에 붙여넣기

---

### 로컬에서 실행하기
```bash
# 1. 필요 패키지 설치
pip install -r requirements.txt

# 2. API 키 환경변수 설정 (선택 — 입력창 대신 사용 가능)
export ANTHROPIC_API_KEY=sk-ant-...

# 3. 앱 실행
streamlit run app.py
```
브라우저에서 `http://localhost:8501` 열기

---

### Streamlit Cloud 무료 배포 (어디서든 접속)
1. GitHub에 이 폴더 업로드
2. [share.streamlit.io](https://share.streamlit.io) 에서 저장소 연결
3. **Secrets** 에 `ANTHROPIC_API_KEY = "sk-ant-..."` 추가
4. 배포 완료 → 링크 공유 가능
    """)
