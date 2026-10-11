# soundable

PDF에서 뽑은 텍스트를, 소리로 읽어 줬을 때 듣기 좋게 고쳐서 Word 문서(docx)로 저장하는 도구다.
시각장애인이 쓴다.

PDF에서 텍스트를 그대로 뽑으면 줄이 바뀌는 자리마다 문장이 끊기고 단어 사이에 엉뚱한 띄어쓰기가 낀다.
읽어 주는 프로그램은 그것을 그대로 읽는다. soundable은 그런 자리를 찾아 고친다.

- 2023년에 만들었다. 커밋은 2023-06-16 ~ 07-15 사이 17개다
- 처음에는 점자(BRF) 변환을 목표로 했다가 소리 쪽으로 방향을 바꿨다
- 소스 첫 줄에 적힌 이름은 `text correction : version0`이다

## 구성

| 경로 | 무엇 |
|---|---|
| `source/text_correction.py` | 프로그램 전체. 파일 하나다 |
| `logo.jpg` | 시작 화면에 띄우는 로고 |
| `executable/Windows11/dist.zip` | Windows 실행파일. 안에 `dist/text_correction/text_correction.exe`가 있다. x86-64, Python 3.11로 묶었다 |
| `executable/macOS_arm64/soundable.zip` | macOS 앱. 안에 `dist/text_correction.app`이 있다. Apple Silicon(arm64), Python 3.10으로 묶었다 |

실행파일은 둘 다 파일 하나가 아니라 폴더째 쓰는 배포본이다. 압축을 풀어 그 안에서 실행한다.
`executable/Windows`와 `source/empty`는 내용이 없는 파일이다.

## 쓰는 법

실행하면 로고가 5초 뜨고, 버튼 두 개가 있는 창이 열린다.

| 버튼 | 하는 일 |
|---|---|
| pdf 불러오기 | PDF를 고르면 그 PDF가 있는 폴더에 같은 이름의 `.docx`가 생긴다. 같은 이름이 이미 있으면 `_1`, `_2`를 붙인다 |
| text 직접 입력하기 | 입력 창에 글을 넣고 제출하면 `result.docx`가 생긴다. 소스로 실행했을 때는 `text_correction.py`가 있는 폴더에 저장된다 |

저장이 끝나면 「변환 완료!!!」 창이 3초 떴다가 닫힌다.
고친 자리 뒤에는 `(고침)`이라는 표시가 붙는다.

## 처리 흐름

PDF를 불러오면 이 순서로 돈다.

1. `pdfminer.six`의 `extract_text`로 텍스트를 뽑는다
2. 기호를 바꾼다 (`correction_korean_special`)
3. 끊긴 줄을 이어 문단을 만든다 (`save_file_docx`)
4. `python-docx`로 문단마다 단락 하나씩 넣어 저장한다

텍스트를 직접 넣었을 때는 2번 뒤에 띄어쓰기 교정(`correction_korean_space`)이 한 번 더 돈다.
PDF 경로에서는 띄어쓰기 교정을 돌리지 않는다.

**기호 바꾸기.** 소스 주석에는 「교육 지원 인력 자료작성 요령(2022)」의 규칙 3·10·12·13·14·15·22번이라고 적혀 있다.

- 로마 숫자 문자(Ⅰ, Ⅱ, ⅲ …)를 알파벳(I, II, iii …)으로 바꾼다
- `※` → `*`, `→`·`⇒` → `->`, `…` → `...`
- `×` → `*`, `÷` → `//`, `±` → `플러스 마이너스`, `㎡` → `제곱미터`, `㎝` → `cm`. `+`와 `=`는 앞뒤를 띄운다
- 전각 문장부호(．，！？＂＇)를 반각으로, 탭을 공백 네 칸으로 바꾼다
- 반각 문장부호 앞에 붙은 공백을 지운다

**줄 잇기.** 빈 줄은 버린다. 마침표로 끝나지 않은 줄은 다음 줄과 붙인다.
처음 10줄과 18자가 안 되는 줄은 제목으로 보고 붙이지 않는다.

**띄어쓰기 교정.**

- 입력 글에서 7번 이상 나온 단어를 모은다. 그 단어 가운데에 공백이 낀 모양(`단 어`)을 찾아 붙인다. 영문·숫자로만 된 단어는 뺀다
- 혼자 떨어진 조사(` 은 `, ` 를 `, ` 에서 ` …)를 앞말에 붙인다
- 혼자 떨어진 문장 끝(` 다. `, ` 이다. `, ` 한다. ` …)을 앞말에 붙인다
- 접속어(그래서, 그러나, 따라서 …) 가운데 낀 공백을 없앤다

## 소스로 실행

```
pip install pdfminer.six python-docx screeninfo Pillow
python source/text_correction.py
```

tkinter가 있어야 한다.
로고는 `text_correction.py`와 같은 폴더에서 찾는다. 저장소 배치 그대로면 못 찾고, 로고 없이 실행된다.

실행파일은 PyInstaller로 묶는다. 소스 주석에 적힌 명령은 이렇다.

```
pyinstaller --noconsole --add-data "logo.jpg:." text_correction.py    # macOS
pyinstaller --noconsole --add-data "logo.jpg;." text_correction.py    # Windows
```

## 기술

Python, tkinter, pdfminer.six, python-docx, screeninfo, Pillow, PyInstaller

PyPDF2가 아니라 pdfminer.six를 쓴 까닭은 소스 주석에 있다.
PDF 하나에 인코딩이 여럿 섞여 있어도 pdfminer.six는 알아서 가려낸다. PyPDF2는 그러지 못했다.

## 알아둘 것

- 붙이면 안 되는 자리를 붙이기도 한다
- 인코딩이 섞인 PDF에서는 글자가 깨질 수 있다
- 텍스트 순서가 뒤섞여 저장된 PDF는 바로잡지 못한다
- 창 안의 글자를 선택할 수 없다. 시각장애인이 쓰려면 선택할 수 있어야 한다
- OCR은 계획에만 있었고 들어가지 않았다. 글자가 그림으로 들어간 PDF는 다루지 못한다
