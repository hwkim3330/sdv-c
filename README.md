# SDV-C (Software-Defined Vehicle Concepts)

## 📚 프로젝트 개요

이 저장소는 SDV(Software-Defined Vehicle) 관련 개념과 표준화 동향을 분석하고 문서화하는 프로젝트입니다. 중국, 독일, 일본의 SDV 표준화 동향과 지능형 연결 차량 서비스 인터페이스 사양을 포함합니다.

## 📁 포함된 문서

### 1. SDV 개념 및 표준화 동향 (한국어)
- **파일명**: `TalkFile_SDV 개념 및 중독일 표준화 동향_최동근작성.pdf.pdf`
- **작성자**: 최동근
- **내용**: SDV의 기본 개념과 중국, 독일, 일본의 표준화 동향 분석

### 2. SDV Intelligent Connected Vehicle Service Interface Specification (중국어)

#### Part 1: Atomic Service API Interface
- **파일명**: `SDV Intelligent Connected Vehicle Service Interface Specification Part 1 Atomic Service API Interface Version 4 Beta 1(중국어).pdf`
- **버전**: Version 4 Beta 1
- **내용**: 원자 서비스 API 인터페이스 사양

#### Part 2: Device Abstraction API Interface  
- **파일명**: `SDV Intelligent Connected Vehicle Service Interface Specification Part 2 Device Abstraction API Interface Version 4 Beta 1(중국어).pdf`
- **버전**: Version 4 Beta 1
- **내용**: 디바이스 추상화 API 인터페이스 사양

### 3. 중국 SDV 표준 소개 (KETI)
- **파일명**: `중국SDV표준 소개_KETI 박부식0826.pptx`
- **작성자**: KETI 박부식
- **내용**: 중국 SDV 표준화 전략과 API 구조 분석 (23 슬라이드)

## 🛠️ 설치 및 실행

### 필요 조건
- Python 3.8 이상
- pip 패키지 관리자

### 설치 방법

1. 저장소 클론
```bash
git clone https://github.com/hwkim3330/sdv-c
cd sdv-c
```

2. 가상환경 생성 및 활성화
```bash
python3 -m venv venv
source venv/bin/activate  # Linux/Mac
# 또는
venv\Scripts\activate  # Windows
```

3. 필요 패키지 설치
```bash
pip install pypdf2 python-pptx
```

### 프레젠테이션 생성

PDF 문서를 분석하여 PowerPoint 프레젠테이션을 생성합니다:

```bash
python create_presentation.py
```

생성된 `SDV_Presentation.pptx` 파일을 확인하세요.

## 📊 생성된 프레젠테이션 파일

### 자동 생성 프레젠테이션

1. **SDV_Presentation.pptx**
   - PDF 문서 자동 분석 결과
   - 목차별 구조화

2. **중국SDV표준_소개_KETI_박부식0826_수정본.pptx**
   - KETI 원본 개선판
   - Executive Summary 추가
   - 한국 산업 대응 전략 포함

3. **SDV_종합분석_보고서.pptx** (28 슬라이드)
   - 완전한 SDV 분석 보고서
   - Part 1: SDV 개념 및 시장 전망
   - Part 2: 글로벌 표준화 현황
   - Part 3: 중국 SDV 표준 심층 분석  
   - Part 4: 기술 비교 분석
   - Part 5: 한국 대응 전략
   - 예산 계획 및 성공 지표

4. **SDV_기술심화_분석.pptx**
   - 기술 아키텍처 상세
   - API 설계 패턴
   - 보안 아키텍처

## 🚀 주요 기능

- **PDF 텍스트 추출**: PyPDF2를 사용한 PDF 문서 텍스트 추출
- **자동 프레젠테이션 생성**: python-pptx를 사용한 PowerPoint 파일 생성
- **다국어 문서 처리**: 한국어 및 중국어 문서 지원

## 📝 SDV 핵심 개념

### Software-Defined Vehicle (SDV)
- 소프트웨어 중심의 차량 아키텍처
- 하드웨어와 소프트웨어의 분리
- OTA(Over-The-Air) 업데이트 지원
- 서비스 지향 아키텍처(SOA)

### 표준화 동향
- **중국**: 지능형 연결 차량 서비스 인터페이스 표준
- **독일**: AUTOSAR Adaptive Platform
- **일본**: 차량 소프트웨어 플랫폼 표준화

## 🔧 기술 스택

- **Python 3.x**: 메인 프로그래밍 언어
- **PyPDF2**: PDF 파일 처리
- **python-pptx**: PowerPoint 프레젠테이션 생성

## 📈 향후 계획

- [ ] 중국어 문서 자동 번역 기능 추가
- [ ] 더 정교한 콘텐츠 추출 알고리즘 개발
- [ ] 웹 기반 뷰어 개발
- [ ] 추가 SDV 관련 문서 수집 및 분석

## 🤝 기여하기

이 프로젝트에 기여하고 싶으시면:

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📄 라이센스

이 프로젝트는 MIT 라이센스 하에 배포됩니다.

## 👥 연락처

프로젝트 관리자: [@hwkim3330](https://github.com/hwkim3330)

프로젝트 링크: [https://github.com/hwkim3330/sdv-c](https://github.com/hwkim3330/sdv-c)

---

**Note**: 중국어 PDF 문서는 텍스트 추출에 제한이 있을 수 있습니다. 더 나은 처리를 위해 OCR 기능을 추가하는 것을 고려하고 있습니다.