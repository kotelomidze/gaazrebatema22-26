/**
 * საგამოცდო ნაშრომების მიმღები — Google Apps Script.
 *
 * გაშვება (ერთხელ):
 *   1. script.google.com → New project
 *   2. წაშალე იქ არსებული კოდი და ჩასვი ეს ფაილი მთლიანად
 *   3. Deploy → New deployment → ტიპი: Web app
 *        Execute as:      Me
 *        Who has access:  Anyone
 *   4. დააკოპირე მიღებული URL და ჩასვი config.json-ის webhookUrl ველში
 *
 * თუ ცხრილშიც გინდა ჩანაწერების დაგროვება, შექმენი Google Sheet,
 * მისამართიდან ამოიღე ID და ჩასვი SHEET_ID-ში. სურვილისამებრია.
 */

var RECIPIENT = 'lomidze.kote1@gmail.com';
var SHEET_ID = '';

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var name = ((data.student && data.student.firstName || '') + ' ' +
                (data.student && data.student.lastName || '')).trim() || 'უცნობი მოსწავლე';

    var answers = data.comprehensionAnswers || {};
    var key = data.correctAnswers || {};
    var numbers = Object.keys(answers).concat(Object.keys(key))
      .filter(function (v, i, a) { return a.indexOf(v) === i; })
      .sort(function (a, b) { return Number(a) - Number(b); });

    var correct = 0;
    var keyKnown = 0;
    var lines = [];

    numbers.forEach(function (n) {
      var given = answers[n] || '—';
      var right = key[n] || '';
      var mark = '';
      if (right) {
        keyKnown++;
        if (given === right) { correct++; mark = '✓'; }
        else { mark = '✗  (სწორი: ' + right + ')'; }
      }
      lines.push('  ' + n + '. ' + given + '  ' + mark);
    });

    var score = (data.comprehensionTotal != null) ? data.comprehensionTotal : correct;
    var maxScore = data.comprehensionMaxPoints || numbers.length;

    var body = [];
    body.push('მოსწავლე: ' + name);
    body.push('ტექსტი: ' + (data.testTitle || data.testId || '—'));
    body.push('გაგზავნის დრო: ' + formatDate(data.submittedAt));
    body.push('');
    body.push('──────────────────────────────');
    body.push('1. ტექსტის გააზრება');
    body.push('──────────────────────────────');

    if (keyKnown === numbers.length && numbers.length > 0) {
      body.push('ქულა: ' + score + ' / ' + maxScore);
    } else if (keyKnown > 0) {
      body.push('ქულა: ' + correct + ' / ' + keyKnown + '  (ნაწილობრივი — გასაღები არასრულია)');
    } else {
      body.push('ქულა ვერ დაითვალა — სწორი პასუხები შევსებული არ არის');
    }
    body.push('');
    body.push(lines.length ? lines.join('\n') : '  (პასუხები არ არის მონიშნული)');

    body.push('');
    body.push('──────────────────────────────');
    body.push('2. თხზულება — ' + (data.essayWordCount || 0) + ' სიტყვა');
    body.push('──────────────────────────────');
    body.push('');
    body.push(data.essayText || '(თხზულება არ დაწერილა)');

    MailApp.sendEmail({
      to: RECIPIENT,
      subject: 'ნაშრომი: ' + name + ' — ' + (data.testTitle || data.testId || ''),
      body: body.join('\n')
    });

    if (SHEET_ID) {
      var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(['თარიღი', 'მოსწავლე', 'ტექსტი', 'ქულა', 'მაქსიმუმი', 'სიტყვა', 'თხზულება']);
      }
      sheet.appendRow([
        new Date(), name, data.testTitle || data.testId,
        score, maxScore, data.essayWordCount || 0, data.essayText || ''
      ]);
    }

    return json({ status: 'success' });

  } catch (error) {
    // შეცდომაც იმეილით მოვა, რომ ნაშრომი უკვალოდ არ დაიკარგოს
    try {
      MailApp.sendEmail(RECIPIENT, 'ნაშრომის მიღების შეცდომა',
        String(error) + '\n\n' + (e && e.postData ? e.postData.contents : '(მონაცემები არ არის)'));
    } catch (ignored) {}
    return json({ status: 'error', error: String(error) });
  }
}

function formatDate(iso) {
  try {
    return Utilities.formatDate(new Date(iso), 'Asia/Tbilisi', 'dd.MM.yyyy HH:mm');
  } catch (e) {
    return iso || '';
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * შესამოწმებლად: გაუშვი ეს ფუნქცია რედაქტორიდან (Run ღილაკი).
 * პირველად Google ნებართვას მოგთხოვს. თუ იმეილი მოვიდა, ყველაფერი წესრიგშია.
 */
function testEmail() {
  doPost({
    postData: {
      contents: JSON.stringify({
        student: { firstName: 'ტესტი', lastName: 'ტესტიშვილი' },
        testTitle: 'შოთა რუსთაველი – „ვეფხისტყაოსანი“',
        submittedAt: new Date().toISOString(),
        comprehensionAnswers: { 2: 'გ', 3: 'ა', 4: 'ბ' },
        correctAnswers: { 2: 'გ', 3: 'ა', 4: 'დ' },
        comprehensionTotal: 2,
        comprehensionMaxPoints: 10,
        essayText: 'ეს სატესტო თხზულებაა.',
        essayWordCount: 3
      })
    }
  });
}
