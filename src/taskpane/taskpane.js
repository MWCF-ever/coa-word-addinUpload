
// 捕获运行时错误
window.onerror = function(message, source, lineno, colno, error) {
  const log = `[window.onerror] ${message} at ${source}:${lineno}:${colno}`;
  logToPage(log);
};

// 捕获未处理的 Promise 错误
window.addEventListener("unhandledrejection", function(event) {
  const log = `[unhandledrejection] ${event.reason}`;
  logToPage(log);
});

function logToPage(message) {
  console.log(message);
  const logContainer = document.getElementById("log");
  if (logContainer) {
    const logEntry = document.createElement("div");
    logEntry.textContent = message;
    logContainer.appendChild(logEntry);
  }
}


Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    logToPage("Office.js 加载成功");
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
  }).catch((err) => {
  logToPage("Office 加载失败: " + err.message);
});

export async function run() {
  return Word.run(async (context) => {
    // Insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

    // Change the paragraph color to blue.
    paragraph.font.color = "blue";

    await context.sync();
  });
}
