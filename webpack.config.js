const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");

module.exports = {
  // 1. 진입점 설정: taskpane.js를 시작점으로 사용합니다.
  entry: "./taskpane.js",

  // 2. 플러그인 설정: taskpane.html을 템플릿으로 사용하여 빌드 결과물을 만듭니다.
  plugins: [
    new HtmlWebpackPlugin({
      template: "./taskpane.html",
      filename: "taskpane.html"
    })
  ],

  // 3. 개발 서버 설정: Office 추가기능 실행을 위해 HTTPS 보안 연결이 필수입니다.
  devServer: {
    port: 3000,
    // [수정] https: true 대신 최신 server 옵션을 사용합니다.
    server: 'https', 
    // Office 앱에서 로컬 서버의 리소스를 불러올 수 있도록 CORS 헤더를 허용합니다.
    headers: {
      "Access-Control-Allow-Origin": "*"
    },
    // 파일 변경 시 브라우저를 자동으로 새로고침합니다.
    hot: true
  },

  // 4. 모듈 해석: 확장자 생략 가능 설정
  resolve: {
    extensions: [".js", ".html"]
  }
};