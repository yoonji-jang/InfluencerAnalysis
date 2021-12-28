# InfluencerAnalysis
#Usage
1. input.txt 에 input 값을 입력한다. 
	DEVELOPER_KEY : Data API 의 service key
	INFLUENCER_SHEET : a,b 동작을 할 sheet number
	VIDEO_SHEET : c 동작을 할 sheet number
	START_COL : A=1, B=2, ...
	MAX_RESULT : c 동작에서 N개의 video 에 대한 정보 수집
2. excel 의 format 은 그대로 사용하고 sheet 2의 video id 열 (B5~ Bn) 에 video id를 추가한다.
3. InfluencerAnalysis.exe 를 실행한다. (in console 또는 더블클릭)
	프로그램 실행 시 a,b,c 전부 수행된다.
4. 프로그램이 종료된 후 output file 을 확인한다. 


----------------------------------------------------------------
1. 기간 : 이번달 내 개발 완료로 목표 (~2021.12.31)
2. 상세 내용
- Input
a. 유튜브프로필ID (예 : UC-Bsa2ivAGWq7bsSPrPGFVA)
b. 최근n개 post (예 : 30)
c. 동영상ID (예 : 2CbF0CIdTiQ)
- Output
a-1. 인플루언서 유튜브 프로필 url
a-2. 인플루언서 유튜브 프로필 이미지
a-3. 채널 구독자 수
b-1. 최근 n개 post의 누적 조회수
b-2. 최근 n개 post의 누적 좋아요수
b-3. 최근 n개 post의 누적 댓글수
b-4. 최근 n개 post의 Engagement rate ( (좋아요수+댓글수)/(조회수) )
(not supported)
(a-4. Age/Gender Rate)
(a-5. 구독자 지역분포)
(a-6. 구독자 언어분포)


c-1. 특정 영상의 url
c-2. 특정 영상의 full 제목
c-3. 특정 영상의 조회수
c-4. 특정 영상의 좋아요수
c-5. 특정 영상의 댓글수
c-6. 특정 영상의 Thumbnail

