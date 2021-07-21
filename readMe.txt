service 실행 ***********************************************************************************
1) Webtest_5G접속
2) cmd창 실행
3) ipconfig (서버ip 확인)
4) D: CD server\KTiFrame (서비스 경로로 이동)
  D: CD D:\kt-automation-testing-team_ykkwon\nodeJS\KTiFrame
5) pm2 start server.js --watch --no-daemon (서버 실행)
6)  service 확인 
    브라우저에서 http://서버ip:3000  주의)port:3000
   ex) http://192.168.1.6:3000

server setting ***********************************************************************************
1. nodeJS 설치
2. pm2 설치
   npm install pm2 -g