/**
 * Жаңалық жарияланған бойда GitHub-қа «превью беттерін қазір жаса» деген белгі жібереді.
 * Сонда /n/<id>/ беті 1–2 минутта дайын болады да, сілтеме WhatsApp/Telegram-да
 * суретімен көрінеді (бұрын GitHub-тың кестесін бірнеше сағат күтетін).
 *
 * ОРНАТУ (бір рет):
 * 1. GitHub → Settings → Developer settings → Fine-grained tokens → Generate new token.
 *    Repository access: тек byNurlykhan/ulandigital.  Permissions → Contents: Read and write.
 * 2. «Ұлан — База» кестесі → Кеңейтімдер → Apps Script.
 *    Жоба параметрлері (⚙) → Script properties → GITHUB_TOKEN = <токен>.
 * 3. Осы файлды сол жобаға жаңа файл етіп қос.
 * 4. Жаңалықты сақтайтын функцияның соңына (жаңалық кестеге жазылған соң) мына жолды қос:
 *        triggerNewsPages_();
 * 5. «Тексеру» үшін редактордан testTriggerNewsPages функциясын бір рет іске қос
 *    (алғаш рет Google рұқсат сұрайды – «Allow»).
 *    GitHub → Actions → «Жаңалық превью беттері» жаңа іске қосылуы көрінуі керек.
 */
function triggerNewsPages_() {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) { console.warn('GITHUB_TOKEN орнатылмаған – превью беттері кестемен ғана жаңарады.'); return false; }
  try {
    const res = UrlFetchApp.fetch('https://api.github.com/repos/byNurlykhan/ulandigital/dispatches', {
      method: 'post',
      contentType: 'application/json',
      headers: {
        Authorization: 'Bearer ' + token,
        Accept: 'application/vnd.github+json',
        'X-GitHub-Api-Version': '2022-11-28'
      },
      payload: JSON.stringify({ event_type: 'news-published' }),
      muteHttpExceptions: true
    });
    const ok = res.getResponseCode() === 204;
    if (!ok) console.warn('GitHub жауабы: ' + res.getResponseCode() + ' ' + res.getContentText());
    return ok;
  } catch (e) {
    // Жаңалықты жариялауға ешқашан кедергі жасамаймыз
    console.warn('GitHub-қа белгі жіберілмеді: ' + e);
    return false;
  }
}

function testTriggerNewsPages() {
  console.log(triggerNewsPages_() ? 'Жіберілді ✔ – GitHub Actions бетін тексер.' : 'Жіберілмеді – журналды қара.');
}
