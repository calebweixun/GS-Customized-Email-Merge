function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("📧 郵件自動化")
    .addItem("選擇草稿並套印", "showDraftPicker")
    .addToUi();
}

function showDraftPicker() {
  var html = HtmlService.createHtmlOutputFromFile("DraftPicker")
    .setWidth(1000)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, "請選擇範本草稿");
}

function getDraftSubjects() {
  try {
    var threads = GmailApp.search("is:draft", 0, 20);
    return threads.map(function (t) {
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

  // --- 1. 動態偵測標頭位置 ---
  var headers = data[0];
  var emailIdx = -1;
  var resultIdx = -1;
  var tagColumns = []; // 儲存所有需要替換的標籤資訊

  headers.forEach(function (header, index) {
    if (!header) return;
    var cleanHeader = header.toString().trim().toLowerCase();

    // 偵測收件者 Email 欄位 (支援 {{email}} 或 email)
    if (cleanHeader === "{{email}}" || cleanHeader === "email") {
      emailIdx = index;
    }
    // 偵測結果輸出欄位 {{result}}
    else if (cleanHeader === "{{result}}") {
      resultIdx = index;
    }
    // 儲存其他所有標籤 (如 {{name}}, {{link}} 等)
    if (header.toString().includes("{{") && header.toString().includes("}}")) {
      tagColumns.push({
        name: header.toString().trim(),
        index: index,
      });
    }
  });

  // 安全檢查
  if (emailIdx === -1)
    throw new Error("找不到標頭為 '{{email}}' 或 'Email' 的欄位。");
  if (resultIdx === -1)
    throw new Error("找不到標頭為 '{{result}}' 的輸出欄位。");

  var query = 'is:draft subject:"' + selectedSubject + '"';
  var threads = GmailApp.search(query, 0, 1);
  if (threads.length === 0) throw new Error("找不到選定的草稿範本。");

  var template = threads[0].getMessages()[0];
  var htmlBody = template.getBody();
  var originalSubject = template.getSubject();
  var attachments = template.getAttachments();

  // 預定義特殊連結正則 (處理 Gmail 自動補上的 http://)
  var linkComplexRegex =
    /(?:https?:\/\/)?(?:%7B%7B|{{)\s*link\s*(?:%7D%7D|}})(?:\/)?/gi;
  var rawLinkRegex =
    /(?:%7B%7B|{{|\[\[|__)\s*rawlink\s*(?:%7D%7D|}}|\]\]|__)/gi;

  var limit = mode === "preview" ? 2 : data.length;
  var previewResult = null;
  var count = 0;

  // --- 2. 逐列處理資料 ---
  for (var j = 1; j < limit; j++) {
    var recipientEmail = data[j][emailIdx] || "";
    if (!recipientEmail && mode !== "preview") continue;

    var finalSubject = originalSubject;
    var finalHtmlBody = htmlBody;

    // --- 3. 動態標籤替換 ---
    tagColumns.forEach(function (tag) {
      var cellValue = data[j][tag.index] || "";
      var tagName = tag.name.toLowerCase();

      // A. 特殊處理 {{link}} (支援超連結嵌入)
      if (tagName === "{{link}}") {
        finalSubject = finalSubject.replace(linkComplexRegex, cellValue);
        finalHtmlBody = finalHtmlBody.replace(
          linkComplexRegex,
          function (match, offset, fullText) {
            var prefix = fullText.substring(offset - 10, offset);
            if (/href=["']$/i.test(prefix)) {
              return cellValue;
            } else {
              return '<a href="' + cellValue + '">' + cellValue + "</a>";
            }
          },
        );
      }
      // B. 特殊處理 {{rawlink}} (純文字網址)
      else if (tagName === "{{rawlink}}") {
        finalSubject = finalSubject.replace(rawLinkRegex, cellValue);
        finalHtmlBody = finalHtmlBody.replace(rawLinkRegex, cellValue);
      }
      // C. 一般標籤替換 (不分大小寫)
      else {
        var tagRegex = new RegExp(escapeRegExp(tag.name), "gi");
        finalSubject = finalSubject.replace(tagRegex, cellValue);
        finalHtmlBody = finalHtmlBody.replace(tagRegex, cellValue);
      }
    });

    // 4. 轉譯 Emoji
    finalHtmlBody = toSafeHtml(finalHtmlBody);

    if (mode === "preview") {
      previewResult = {
        to: recipientEmail || "範例收件者",
        subject: finalSubject,
        body: finalHtmlBody,
      };
      break;
    }

    var timestamp = Utilities.formatDate(
      new Date(),
      "GMT+8",
      "yyyy-MM-dd HH:mm:ss",
    );

    if (mode === "send") {
      GmailApp.sendEmail(recipientEmail, finalSubject, "", {
        htmlBody: finalHtmlBody,
        attachments: attachments,
      });
      sheet
        .getRange(j + 1, resultIdx + 1)
        .setValue("✅ 已寄出 (" + timestamp + ")");
      count++;
    } else if (mode === "draft") {
      GmailApp.createDraft(recipientEmail, finalSubject, "", {
        htmlBody: finalHtmlBody,
        attachments: attachments,
      });
      sheet
        .getRange(j + 1, resultIdx + 1)
        .setValue("📝 已建草稿 (" + timestamp + ")");
      count++;
    }
  }

  if (mode === "preview") return previewResult;
  return "操作成功！已完成 " + count + " 封郵件處理。";
}

// 輔助函式
function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function toSafeHtml(str) {
  if (!str) return "";
  return Array.from(str)
    .map(function (char) {
      var code = char.codePointAt(0);
      return code > 127 ? "&#" + code + ";" : char;
    })
    .join("");
}
