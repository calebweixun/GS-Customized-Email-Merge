function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // 建立選單
  ui.createMenu('📧 郵件自動化')
      .addItem('選擇草稿並套印', 'showDraftPicker')
      .addToUi();
}

function showDraftPicker() {
  var html = HtmlService.createHtmlOutputFromFile('DraftPicker')
      .setWidth(1000)  // 調至 Google 視窗最大寬度
      .setHeight(700); 
  SpreadsheetApp.getUi().showModalDialog(html, '請選擇範本草稿');
}

function getDraftSubjects() {
  try {
    // 只搜尋最近的 20 個草稿執行緒 (Thread)
    var threads = GmailApp.search("is:draft", 0, 20);
    return threads.map(function(t) {
      return t.getFirstMessageSubject() || "(無主旨)";
    });
  } catch (e) {
    return ["讀取失敗: " + e.message];
  }
}

function processSelectedDraft(selectedSubject, mode) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  var query = 'is:draft subject:"' + selectedSubject + '"';
  var threads = GmailApp.search(query, 0, 1);
  if (threads.length === 0) throw new Error("找不到選定的草稿範本。");
  
  var template = threads[0].getMessages()[0];
  var htmlBody = template.getBody();
  var originalSubject = template.getSubject();
  var attachments = template.getAttachments();

  var nameRegex = /{{\s*(姓名|Name)\s*}}/gi;
  var companyRegex = /{{\s*(公司|Company)\s*}}/gi;

  var limit = (mode === 'preview') ? 2 : data.length;
  var previewResult = null;
  var count = 0;

  for (var j = 1; j < limit; j++) {
    var companyName = data[j][0] || "";   
    var recipientName = data[j][1] || ""; 
    var recipientEmail = data[j][2] || ""; 
    
    if (!recipientEmail && mode !== 'preview') continue;

    // 1. 執行套印替換
    var finalSubject = originalSubject.replace(nameRegex, recipientName).replace(companyRegex, companyName);
    var finalHtmlBody = htmlBody.replace(nameRegex, recipientName).replace(companyRegex, companyName);

    // 2. 轉譯內容中的 Emoji (主旨不轉譯)
    finalHtmlBody = toSafeHtml(finalHtmlBody);

    if (mode === 'preview') {
      previewResult = { to: recipientEmail || "範例收件者", subject: finalSubject, body: finalHtmlBody };
      break; 
    } 

    // 【新增功能】：準備時間戳記
    var timestamp = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss");

    if (mode === 'send') {
      // 直接寄出
      GmailApp.sendEmail(recipientEmail, finalSubject, "", {
        htmlBody: finalHtmlBody,
        attachments: attachments
      });
      // 在第 4 欄 (D) 寫入記錄
      sheet.getRange(j + 1, 4).setValue("✅ 已寄出 (" + timestamp + ")");
      count++;
    } else if (mode === 'draft') {
      // 產生草稿
      GmailApp.createDraft(recipientEmail, finalSubject, "", {
        htmlBody: finalHtmlBody,
        attachments: attachments
      });
      // 在第 4 欄 (D) 寫入記錄
      sheet.getRange(j + 1, 4).setValue("📝 已建草稿 (" + timestamp + ")");
      count++;
    }
  }

  if (mode === 'preview') return previewResult;
  return "操作成功！已完成 " + count + " 封郵件處理 (" + (mode === 'send' ? '直接寄出' : '產生草稿') + ")。";
}

function toSafeHtml(str) {
  if (!str) return "";
  return Array.from(str).map(function(char) {
    var code = char.codePointAt(0);
    return code > 127 ? "&#" + code + ";" : char;
  }).join("");
}
