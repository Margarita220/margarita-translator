from flask import Flask, request, jsonify, Response
import json
import os

os.environ["translators_default_region"] = "CN"

app = Flask(__name__)

def load_dictionaries():
    dict_path = "dictionary.xlsx"
    cn_to_ru = []
    ru_to_cn = []
    
    if os.path.exists(dict_path):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(dict_path, data_only=True)
            sheet = wb.active
            rows = list(sheet.iter_rows(values_only=True))
            if len(rows) > 1:
                headers = [str(h).strip().lower() if h else "" for h in rows]
                if "chinese" in headers and "russian" in headers:
                    cn_idx = headers.index("chinese")
                    ru_idx = headers.index("russian")
                    for row in rows[1:]:
                        cn_val = str(row[cn_idx]).strip().lower() if row[cn_idx] else ""
                        ru_val = str(row[ru_idx]).strip() if row[ru_idx] else ""
                        if cn_val and ru_val:
                            cn_to_ru.append((cn_val, ru_val))
                            ru_to_cn.append((ru_val.lower(), cn_val))
                    cn_to_ru.sort(key=lambda x: len(x), reverse=True)
                    ru_to_cn.sort(key=lambda x: len(x), reverse=True)
        except Exception as e:
            print(f"Ошибка чтения Excel: {e}")
            
    return cn_to_ru, ru_to_cn

CN_TO_RU_LIST, RU_TO_CN_LIST = load_dictionaries()

def fetch_online_translation(text: str, source_lang: str, target_lang: str) -> str:
    import translators as ts
    for engine in ('bing', 'google', 'alibaba', 'sogou'):
        try:
            return ts.translate_text(text, from_language=source_lang, to_language=target_lang, translator=engine)
        except:
            continue
    return "Ошибка сети. Попробуйте еще раз."

def hybrid_translate(text: str, direction: str) -> str:
    cleaned_text = text.strip()
    if not cleaned_text:
        return ""
    
    match_clean = cleaned_text.lower().replace("。", "").replace("，", "").replace(".", "").replace(",", "").strip()
    
    if direction == "CN ⇄ RU":
        if "挤" in match_clean and ("水分" in match_clean or "水份" in match_clean):
            return "Мы можем немного снизить для вас цену."
            
        translation = fetch_online_translation(cleaned_text, "zh-CHS", "ru")
        
        if "三角债" in match_clean and ("треугольн" in translation.lower() or "долг" in translation.lower()):
            translation = translation.replace("треугольных долгов", "взаимных неплатежей").replace("треугольного долга", "взаимных неплатежей").replace("треугольные долги", "взаимные неплатежи")
        
        for cn_term, ru_translation in CN_TO_RU_LIST:
            if cn_term in cleaned_text.lower():
                bad_machine_term = fetch_online_translation(cn_term, "zh-CHS", "ru").lower().strip().replace(".", "")
                if bad_machine_term in translation.lower():
                    start_idx = translation.lower().find(bad_machine_term)
                    end_idx = start_idx + len(bad_machine_term)
                    translation = translation[:start_idx] + ru_translation + translation[end_idx:]
        return translation
            
    else:
        if "снизить" in match_clean and "цену" in match_clean:
            return "我们可以给你挤一点水分。"
            
        translation = fetch_online_translation(cleaned_text, "ru", "zh-CHS")
        
        for ru_term, cn_translation in RU_TO_CN_LIST:
            if ru_term in cleaned_text.lower():
                bad_machine_term = fetch_online_translation(ru_term, "ru", "zh-CHS").lower().strip()
                if bad_machine_term in translation.lower():
                    start_idx = translation.lower().find(bad_machine_term)
                    end_idx = start_idx + len(bad_machine_term)
                    translation = translation[:start_idx] + cn_translation + translation[end_idx:]
        return translation
MANIFEST_JSON = json.dumps({
    "name": "Маргарита Переводчик ИИП",
    "short_name": "Переводчик ИИП",
    "start_url": "/",
    "display": "standalone",
    "background_color": "#ffffff",
    "theme_color": "#1976d2",
    "icons": [
        {
            "src": "https://flaticon.com",
            "sizes": "512x512",
            "type": "image/png"
        }
    ]
}, ensure_ascii=False)

HTML_PAGE = """<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="default">
<meta name="apple-mobile-web-app-title" content="Переводчик ИИП">
<link rel="manifest" href="/manifest.json">
<link rel="apple-touch-icon" href="https://flaticon.com">
<title>Маргарита Переводчик ИИП</title>
<style>
  * { box-sizing: border-box; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; }
  body { margin: 0; padding: 16px; background: #f5f7fa; display: flex; justify-content: center; }
  .container { width: 100%; max-width: 480px; background: #fff; padding: 24px 20px; border-radius: 20px; box-shadow: 0 8px 24px rgba(0,0,0,0.06); text-align: center; }
  h1 { font-size: 22px; margin: 0 0 16px; color: #111; }
  .dir-btn { background: #e3f2fd; color: #1565c0; border: none; padding: 10px 20px; border-radius: 20px; font-size: 15px; font-weight: 600; cursor: pointer; margin-bottom: 18px; transition: 0.2s; }
  .dir-btn:active { transform: scale(0.97); }
  textarea { width: 100%; height: 130px; padding: 14px; border: 1.5px solid #d0d7de; border-radius: 14px; font-size: 16px; resize: none; outline: none; transition: border 0.2s; }
  textarea:focus { border-color: #1976d2; }
  textarea[readonly] { background: #f8f9fa; color: #222; }
  .counter { text-align: right; font-size: 12px; color: #888; margin: 4px 0 14px; }
  .actions { display: flex; justify-content: center; align-items: center; gap: 18px; margin-bottom: 12px; }
  .trans-btn { background: #ffee58; color: #000; border: none; padding: 14px 28px; border-radius: 24px; font-size: 16px; font-weight: 600; cursor: pointer; box-shadow: 0 2px 6px rgba(0,0,0,0.1); }
  .trans-btn:active { transform: scale(0.97); }
  .mic-btn { width: 56px; height: 56px; border-radius: 50%; border: none; background: #ff5252; color: #fff; font-size: 24px; cursor: pointer; display: flex; align-items: center; justify-content: center; box-shadow: 0 4px 10px rgba(255,82,82,0.35); transition: 0.2s; }
  .mic-btn.recording { background: #212121; animation: pulse 1.2s infinite; }
  @keyframes pulse { 0% { box-shadow: 0 0 0 0 rgba(255,82,82,0.6); } 70% { box-shadow: 0 0 0 14px rgba(255,82,82,0); } 100% { box-shadow: 0 0 0 0 rgba(255,82,82,0); } }
  .status { font-size: 13px; min-height: 20px; margin-bottom: 12px; color: #555; }
</style>
</head>
<body>
<div class="container">
  <h1>Маргарита Переводчик ИИП</h1>
  <button id="dirBtn" class="dir-btn">⇄ Китайский ⇄ Русский</button>
  
  <textarea id="inputField" maxlength="3000" placeholder="Введите китайский текст (до 3000 знаков)..."></textarea>
  <div id="counter" class="counter">0/3000</div>
  
  <div class="actions">
    <button id="transBtn" class="trans-btn">文A Перевести</button>
    <button id="micBtn" class="mic-btn" title="Голосовой ввод">🎤</button>
  </div>
  
  <div id="status" class="status"></div>
  <textarea id="outputField" readonly placeholder="Результат перевода..."></textarea>
</div>

<script>
let direction = "CN ⇄ RU";
let isRecording = false;
let recognition = null;

const dirBtn = document.getElementById("dirBtn");
const inputField = document.getElementById("inputField");
const outputField = document.getElementById("outputField");
const counter = document.getElementById("counter");
const transBtn = document.getElementById("transBtn");
const micBtn = document.getElementById("micBtn");
const statusEl = document.getElementById("status");

inputField.addEventListener("input", () => {
  counter.textContent = inputField.value.length + "/3000";
});

dirBtn.addEventListener("click", () => {
  if (direction === "CN ⇄ RU") {
    direction = "RU ⇄ CN";
    dirBtn.textContent = "⇄ Русский ⇄ Китайский";
    inputField.placeholder = "Введите русский текст (до 3000 знаков)...";
  } else {
    direction = "CN ⇄ RU";
    dirBtn.textContent = "⇄ Китайский ⇄ Русский";
    inputField.placeholder = "Введите китайский текст (до 3000 знаков)...";
  }
  inputField.value = "";
  outputField.value = "";
  statusEl.textContent = "";
  counter.textContent = "0/3000";
});

async function doTranslate() {
  const text = inputField.value.trim();
  if (!text) {
    outputField.value = "";
    return;
  }
  outputField.value = "Перевод...";
  try {
    const res = await fetch("/translate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text: text, direction: direction })
    });
    const data = await res.json();
    outputField.value = data.translation;
  } catch (e) {
    outputField.value = "Ошибка соединения с сервером.";
  }
}

transBtn.addEventListener("click", doTranslate);

const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;

micBtn.addEventListener("click", () => {
  if (!SpeechRecognition) {
    statusEl.textContent = "❌ Ваш браузер не поддерживает голосовой ввод.";
    statusEl.style.color = "#d32f2f";
    return;
  }

  if (isRecording && recognition) {
    recognition.stop();
    return;
  }

  recognition = new SpeechRecognition();
  recognition.lang = (direction === "CN ⇄ RU") ? "zh-CN" : "ru-RU";
  recognition.interimResults = true;
  recognition.continuous = false;

  recognition.onstart = () => {
    isRecording = true;
    micBtn.classList.add("recording");
    micBtn.textContent = "⏹";
    statusEl.textContent = "🔴 Говорите сейчас...";
    statusEl.style.color = "#d32f2f";
  };

  recognition.onresult = (event) => {
    let transcript = "";
    for (let i = 0; i < event.results.length; i++) {
      transcript += event.results.item(i).item(0).transcript;
    }
    inputField.value = transcript;
    counter.textContent = transcript.length + "/3000";
  };

  recognition.onerror = (event) => {
    statusEl.textContent = "⚠️ Ошибка микрофона: " + event.error;
    statusEl.style.color = "#e65100";
  };

  recognition.onend = () => {
    isRecording = false;
    micBtn.classList.remove("recording");
    micBtn.textContent = "🎤";
    if (inputField.value.trim()) {
      statusEl.textContent = "✅ Речь распознана!";
      statusEl.style.color = "#2e7d32";
      doTranslate();
    } else {
      statusEl.textContent = "Ничего не распознано. Попробуйте еще раз.";
      statusEl.style.color = "#666";
    }
  };

  recognition.start();
});
</script>
</body>
</html>
"""

@app.route("/")
def index():
    return Response(HTML_PAGE, mimetype="text/html; charset=utf-8")

@app.route("/manifest.json")
def manifest():
    return Response(MANIFEST_JSON, mimetype="application/manifest+json; charset=utf-8")

@app.route("/translate", methods=["POST"])
def translate_route():
    data = request.get_json(force=True)
    result = hybrid_translate(data.get("text", ""), data.get("direction", "CN ⇄ RU"))
    return jsonify({"translation": result})
