1. 크롤링을 활용한 대학 시험 대비 프로그램
2. chatgpt api를 활용하여 prompt를 활용해 pdf와 각종 문서들을 읽어들여 read만 할 수 있는 파일로 요약 처리
3. 각기 대학교들의 e-class를 활용하면 됨

이 프로젝트에 .env파일을 추가하고 자신의 chatgpt api와 ms office를 활용하여 로그인 인증 시스템을 거치면됌
.env파일 코드는 다음과 같음
LMS_ID=(your e-class id)
LMS_PASSWORD=(your e-class password)
EMAIL_ID=(your ms office365 id)
EMAIL_PASSWORD=(your ms office365 password)
OPENAI_API_KEY=(your chatgpt api key)
