# 📊 PPT 목차 기반 자동 취합 도구 (PPT Auto-Merger)

이 프로젝트는 마스터 파워포인트(PPT) 파일에서 목차를 추출하여 엑셀 가이드라인을 생성하고, 협업자들이 개별적으로 작성한 PPT 파일 중 **작성률이 100%인 파일만 자동으로 선별하여 하나의 PPT로 병합**해 주는 파이썬 기반의 업무 자동화 스크립트입니다.

## ✨ 주요 기능
1. **목차 추출 및 엑셀 생성:** 원본 PPT의 슬라이드 제목을 인식하여 `순서`, `목차`, `파일명`, `작성률(%)`이 포함된 엑셀 가이드 파일을 자동 생성합니다.
2. **조건부 자동 병합:** 엑셀 파일에 기록된 작성률이 `100`인 개별 PPT 파일들만 순서대로 모아 하나의 최종 파일로 취합합니다.

## 🛠 시스템 요구사항
- **OS:** Windows / macOS / Linux (크로스 플랫폼 지원)
- **Python:** Python 3.8 이상 권장

## 📦 설치 방법

1. 저장소를 클론(Clone)하거나 다운로드합니다.
```bash
git clone [https://github.com/your-username/my-ppt-manager.git](https://github.com/your-username/my-ppt-manager.git)
cd my-ppt-manager


(선택 사항) 가상환경을 생성하고 활성화합니다.

Bash
# Windows
python -m venv venv
venv\Scripts\activate

# macOS/Linux
python3 -m venv venv
source venv/bin/activate

필요한 라이브러리를 설치합니다.

Bash
pip install -r requirements.txt


🚀 사용 방법
1단계: 원본 PPT에서 목차 추출하기
기준이 되는 원본 파일(master_template.pptx)을 프로젝트 폴더에 넣습니다.

main.py를 실행합니다.

Bash
python main.py
메뉴에서 1을 입력하여 목차를 추출합니다.

생성이 완료되면 ppt_merge_guide.xlsx 파일이 생성됩니다.

2단계: 엑셀 파일 업데이트 및 개별 PPT 준비
생성된 ppt_merge_guide.xlsx를 엽니다.

협업 진행 상황에 맞춰 개별_PPT_파일명을 실제 파일명과 맞추고, 완료된 항목의 작성률(%)을 100으로 변경 후 저장합니다.

개별적으로 작성된 조각 PPT 파일들을 individual_ppts/ 폴더 안에 모아둡니다.

3단계: 완성된 파일 병합하기
다시 main.py를 실행합니다.

메뉴에서 2를 입력하여 취합을 시작합니다.

작업이 끝나면 작성률 100%인 파일들만 취합된 final_output.pptx가 생성됩니다.

📂 프로젝트 구조
Plaintext
my-ppt-manager/
├── individual_ppts/       # 개별 조각 PPT 파일들을 모아두는 폴더
├── main.py                # 실행 스크립트 (추출 및 취합 기능)
├── requirements.txt       # 의존성 패키지 목록
├── README.md              # 프로젝트 설명서 (현재 파일)
├── master_template.pptx   # [사용자 준비] 목차를 추출할 원본 파일
├── ppt_merge_guide.xlsx   # [자동 생성] 추출된 목차 및 관리 엑셀
└── final_output.pptx      # [자동 생성] 최종 취합된 결과물

