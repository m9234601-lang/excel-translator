import streamlit as st
import google.generativeai as genai
import xlwings as xw
import os
import time
import json
import pandas as pd

# 1. 페이지 설정 및 디자인
st.set_page_config(page_title="작업 표준서 다국어 번역기", layout="wide")
st.title("🏭 작업 표준서 다국어 번역 시스템 (웹 공유형)")
st.write("엑셀의 사진, 도형, 서식을 100% 원본 그대로 유지하면서 글자만 자동 전문 번역합니다.")

# 2. 웹 배포용 세션 상태 초기화 (각 접속자의 브라우저 메모리에만 임시 저장)
if "saved_api_key" not in st.session_state:
    st.session_state["saved_api_key"] = ""

# 3. 사이드바 설정 (보안 키 및 언어 선택)
st.sidebar.header("🔧 설정")

# 웹 버전은 보안을 위해 매번 빈칸으로 시작하며, 본인의 키를 넣어야 가동됩니다.
api_key = st.sidebar.text_input(
    "Gemini API Key를 입력하세요:", 
    type="password",
    value=st.session_state["saved_api_key"]
)

# 사용자가 키를 입력하는 즉시 세션 메모리에 실시간 반영
if api_key:
    st.session_state["saved_api_key"] = api_key.strip()

languages = ["한국어", "영어", "일본어", "중국어", "베트남어"]
st.sidebar.subheader("🌐 언어 선택")
source_lang = st.sidebar.selectbox("원본 언어", languages, index=1) # 기본값 영어
target_lang = st.sidebar.selectbox("목적 언어", languages, index=0) # 기본값 한국어

if source_lang == target_lang:
    st.sidebar.error("⚠️ 원본 언어와 목적 언어는 달라야 합니다.")

# 4. 메인 화면 - 파일 업로드 섹션
st.header("📂 엑셀 파일 업로드")
uploaded_file = st.file_uploader("번역할 작업 표준서(Excel) 파일을 선택하세요.", type=["xlsx"])

# 5. 스마트 자동 재시도(Auto-Retry) 기능이 내장된 배치 번역 함수
def translate_batch(text_list, source, target, model):
    if not text_list:
        return {}
        
    mapping_rules = ""
    if source == "한국어" and target == "영어":
        mapping_rules = """
        Crucial Manufacturing Terminology Mapping Rules (KO -> EN):
        "공정명" -> "Process Name", "공정번호" -> "Process No.", "퍼지" -> "Purge", 
        "공타" -> "Dry Shot", "초물" -> "First Article", "구분" -> "Division", "시업시" -> "Startup"
        """
    elif source == "영어" and target == "한국어":
        mapping_rules = """
        Crucial Manufacturing Terminology Mapping Rules (EN -> KO):
        "Process Name" or "fair name" -> "공정명", "Process No." -> "공정번호", 
        "Purge" or "fudge" -> "퍼지", "Dry Shot" or "empty hit" -> "공타", 
        "First Article" -> "초물", "Division" or "Category" -> "구분", "Startup" -> "시업시",
        "work standards" -> "작업표준서", "Status of parts used" -> "부품 사용 현황", "Quality check box" -> "품질 체크 박스"
        """
        
    # 고유 숫자 ID 매칭법으로 AI 오지랖(키 번역 오류) 원천 차단
    input_packages = [{"id": i, "text": text} for i, text in enumerate(text_list)]
        
    prompt = f"""
    You are an expert translator specializing in industrial manufacturing, automotive parts production, safety management, and Standard Operating Procedures (SOP).
    Translate the 'text' field of each object from {source} to {target} reflecting this industrial context.
    
    {mapping_rules}
    
    Strict Rules:
    1. Return the results ONLY as a valid JSON array of objects.
    2. Each object MUST contain the exact original 'id' (integer) and