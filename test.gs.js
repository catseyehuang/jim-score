// 在 code.js 中加入這個測試函式
function testGetScore() {
  const mockEvent = {
    parameter: {
      action: 'getScore'
    }
  };
  const result = handleRequest(mockEvent);
  console.log(result.getContent()); // 應該會印出 {"success":true, "score":...}
}