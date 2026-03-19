// --- КОДЫН ЭХЛЭЛ ---

/****************************************************************************************
 * PRODUCT: LOVE NUMEROLOGY REPORT GENERATOR (V2)
 * ENGINE: ROBUST V26.1 (Smart Doc Insertion, Token Limit Safe, UChat, Regex Cleaner)
 ****************************************************************************************/

const CONFIG = {
  VERSION: "v26.2-Love-Numerology-UChat",
  PRODUCT_NAME: "Хайр Дурлалын Зураг Төөрөг: Ирээдүйн Ханийн Нууц Код",
  SHEET_NAME: "Sheet1",
  BATCH_SIZE: 3,
  GEMINI_MODEL: "gemini-2.5-flash",
  TEMPERATURE: 0.35,

  COLUMNS: {
    NAME: 0, ID: 1, INPUT: 2, PDF: 3, STATUS: 4,
    TOKEN: 5, DEBUG: 6, DATE: 7, VER: 8, ERROR: 9
  },

  UCHAT: {
    ENDPOINT: "https://www.uchat.com.au/api/subscriber/send-content",
    DELIVERY_MESSAGE: `Сайн байна уу, {{NAME}}? 🔮\n\nТаны "Хайр Дурлалын Зураг Төөрөг: Ирээдүйн Ханийн Нууц Код" бэлэн боллоо.\n\nДоорх товч дээр дарж татаж авна уу. 👇`,
    DELIVERY_BTN_TEXT: "📥 Тайлан татах"
  },

};

function getProperty(key) {
  const val = PropertiesService.getScriptProperties().getProperty(key);
  if (!val) throw new Error(`MISSING SCRIPT PROPERTY: ${key}`);
  return val;
}

// ==========================================
// 🚀 ROBUST MAIN ENGINE (AUTO-HEALING)
// ==========================================
function main() {
  const START_TIME = new Date().getTime();
  const TIME_LIMIT_MS = 5 * 60 * 1000; // 5.0 minutes

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) throw new Error(`SYSTEM: "${CONFIG.SHEET_NAME}" sheet олдсонгүй.`);
    const rows = sheet.getDataRange().getValues();
    let processedCount = 0;

    const KEYS = {
      GEMINI: getProperty("GEMINI_API_KEY"),
      TEMPLATE: getProperty("TEMPLATE_ID"),
      UCHAT: getProperty("UCHAT_API_KEY"),
      FOLDER: getProperty("FOLDER_ID")
    };

    const COLS = CONFIG.COLUMNS;

    for (let i = 1; i < rows.length; i++) {
      if (processedCount >= CONFIG.BATCH_SIZE) break;

      // Premium Upgrade 2: Time Limit Safeguard
      if (new Date().getTime() - START_TIME > TIME_LIMIT_MS) {
        console.warn("Time limit approaching. Stopping early to prevent Google Apps Script timeout.");
        break;
      }

      const row = rows[i];
      const name = String(row[COLS.NAME] || "Эрхэм");
      const contactID = row[COLS.ID];
      const inputData = String(row[COLS.INPUT]);
      const status = String(row[COLS.STATUS] || "");
      const rawDate = row[COLS.DATE];

      const pdfCell = sheet.getRange(i + 1, COLS.PDF + 1);
      const statusCell = sheet.getRange(i + 1, COLS.STATUS + 1);
      const errorCell = sheet.getRange(i + 1, COLS.ERROR + 1);
      const tokenCell = sheet.getRange(i + 1, COLS.TOKEN + 1);
      const debugCell = sheet.getRange(i + 1, COLS.DEBUG + 1);
      const dateCell = sheet.getRange(i + 1, COLS.DATE + 1);
      const verCell = sheet.getRange(i + 1, COLS.VER + 1);

      if (!inputData) continue;
      if (status === "АМЖИЛТТАЙ" || status.includes("ХЯНАХ ШААРДЛАГАТАЙ") || status.includes("24 цаг хэтэрсэн")) continue;

      let isRetry = false;
      if (status === "Боловсруулж байна...") {
        if (rawDate instanceof Date) {
          const nowMs = new Date().getTime();
          const startMs = rawDate.getTime();
          const diffMinutes = (nowMs - startMs) / (1000 * 60);

          if (diffMinutes > 15) {
             isRetry = true;
             console.log(`Timeout recovery for user. Stuck for ${Math.round(diffMinutes)} mins.`);
          } else {
             continue;
          }
        } else {
           continue;
        }
      }

      statusCell.setValue("Боловсруулж байна...");

      const startTime = new Date();
      dateCell.setValue(Utilities.formatDate(startTime, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"));
      SpreadsheetApp.flush();

      try {
        console.log(`Processing Love Numerology logic for user...`);

        // 1. DATA PREP: AI PARSING (from old_code, adapted)
        const profileData = parseAndCalculateProfile(inputData, KEYS.GEMINI);

        // 3. GENERATE REPORT (The Voice) - Split Calls (adapted)
        const reportResult = generateSequentialReport(profileData, KEYS.GEMINI);

        // 4. SAFE PDF DELIVERY ENGINE
        const pdfUrl = createPdfSafely(profileData.name || "Эрхэм", reportResult.text, KEYS.TEMPLATE, KEYS.FOLDER);

        sendUChatProven(contactID, pdfUrl, profileData.name || "Эрхэм", KEYS.UCHAT);

        const totalTokens = (profileData.parsingUsage || 0) + reportResult.usage;
        const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");

        pdfCell.setValue(pdfUrl);
        statusCell.setValue("АМЖИЛТТАЙ");
        tokenCell.setValue(totalTokens);
        debugCell.setValue(JSON.stringify(profileData));
        dateCell.setValue(now);
        verCell.setValue(CONFIG.VERSION);
        errorCell.setValue("");

        processedCount++;

      } catch (err) {
        let errorMsgStr = err.toString();
        let mongolianError = "Системийн алдаа: " + errorMsgStr;

        if (errorMsgStr.includes("24H_LIMIT") || errorMsgStr.includes("window")) {
            mongolianError = "Фэйсбүүк 24 цаг хэтэрсэн тул мессеж явуулах эрх хаагдсан байна.";
            statusCell.setValue("24 цаг хэтэрсэн");
            errorCell.setValue(mongolianError);
            continue;
        } else if (errorMsgStr.includes("Gemini") || errorMsgStr.includes("JSON Parse")) {
            mongolianError = "AI (Gemini) хариу өгсөнгүй эсвэл түр зуур хэт ачаалалтай байна.";
        } else if (errorMsgStr.includes("UChat token") || errorMsgStr.includes("user_ns")) {
            mongolianError = "UChat тохиргоо эсвэл харилцагчийн код буруу байна.";
        } else if (errorMsgStr.includes("Drive") || errorMsgStr.includes("FOLDER_ID")) {
            mongolianError = "Google Drive лимит хэтэрсэн эсвэл ID буруу байна.";
        }

        console.error(`Error: ${errorMsgStr}`);
        errorCell.setValue(mongolianError);

        if (isRetry || status === "") {
             statusCell.setValue("Дахин оролдож байна (1)");
        } else if (status === "Дахин оролдож байна (1)") {
             statusCell.setValue("Дахин оролдож байна (2)");
        } else if (status === "Дахин оролдож байна (2)") {
             statusCell.setValue("ХЯНАХ ШААРДЛАГАТАЙ");
        } else {
             statusCell.setValue("Дахин оролдож байна (1)");
        }
      }
    }
  } finally {
    lock.releaseLock();
  }
}


// ==========================================
// 2. MATH & NUMEROLOGY ENGINE
// ==========================================
const NUMEROLOGY = {
  LIFE_PATH_MAP: {
    1: { name: "Бие даасан манлайлагч", desc: "Хайранд хүчтэй байр суурьтай, шийдвэр гаргагч." },
    2: { name: "Сэтгэл холбооны бүтээгч", desc: "Гүн ойлголцол, хамт амьдрах зохицолд төвлөрдөг." },
    3: { name: "Илэрхийллийн хайрлагч", desc: "Ярилцлага, урам, сэтгэлээ ил гаргах нь чухал." },
    4: { name: "Тогтвортой суурь бүтээгч", desc: "Харилцаанд найдвартай, бүтэцтэй орон зай шаарддаг." },
    5: { name: "Өөрчлөлтийн эрэлчин", desc: "Эрч хүчтэй, эрх чөлөөтэй харилцааг илүүд үздэг." },
    6: { name: "Гэр бүлсэг халамжлагч", desc: "Халамж, хамгаалалт, гэр бүлийн дулаан эрчимтэй." },
    7: { name: "Дотоод ертөнцийн шинжээч", desc: "Сэтгэлээ нээхэд хугацаа хэрэгтэй, гүн холбоо хайдаг." },
    8: { name: "Хил хязгаарын эзэн", desc: "Үнэ цэнэ, хүндлэл, бодит үр дүнг чухалчилдаг." },
    9: { name: "Өндөр мэдрэмжит өгөөмөр", desc: "Том сэтгэлтэй ч сэтгэл хөдлөлийн хамгаалалт хэрэгтэй." },
    11: { name: "Зөн совинтой хайрлагч", desc: "Сэтгэлийн гүн холбоо, үггүй ойлголцлыг хайдаг мастер код." },
    22: { name: "Амьдралын түншлэл бүтээгч", desc: "Хамтын амьдралыг бодитоор босгох чадвартай." },
    33: { name: "Сэтгэл эмчлэгч", desc: "Хайрыг халамж ба утгатай амьдрал болгон хувиргадаг." }
  },
  MATRIX_EXCESS_MAP: {
    1: { title: "Манлайллын хэт хүч", desc: "Харилцаанд хэт хяналт тогтоох эрсдэлтэй." },
    2: { title: "Мэдрэмжийн ачаалал", desc: "Бусдын сэтгэл хөдлөлийг өөртөө хэт наах хандлагатай." },
    3: { title: "Илэрхийллийн хэт ачаалал", desc: "Сэтгэлээ олон сувгаар тарааж, тогтворгүйдэх эрсдэлтэй." },
    4: { title: "Хатуу хяналт", desc: "Уян хатан байдлыг багасгаж, харилцааг чангалж болзошгүй." },
    5: { title: "Тогтворгүй эрчим", desc: "Шинэ мэдрэмж хөөхөөс болж харилцаа тогтворгүйдэх магадлалтай." },
    6: { title: "Өөрийгөө золиослох", desc: "Хэт халамжлаад өөрийн орон зайгаа алдах эрсдэлтэй." },
    7: { title: "Дотогш хаагдах", desc: "Сэтгэлээ дотроо хадгалж, ойлголцлыг таслах магадлалтай." },
    8: { title: "Хатуу шаардлага", desc: "Хил хязгаар зөв ч хэт хатуу бол хайр хөрнө." },
    9: { title: "Өнгөрсөнд гацах", desc: "Дурсамж, гомдол удаан хадгалах эрсдэлтэй." }
  },
  MISSING_NUMBER_MAP: {
    1: { title: "Өөрийгөө илэрхийлэх цоорхой", risk: "Хүсэл хэрэгцээгээ тод хэлж чадалгүй буруу ойлголцол үүсгэх." },
    2: { title: "Ойлголцлын цоорхой", risk: "Хосын нарийн мэдрэмжийг ойлгоход саадтай." },
    3: { title: "Ярианы цоорхой", risk: "Ярилцлага дутагдаж сэтгэл холдох." },
    4: { title: "Тогтвортой байдлын цоорхой", risk: "Харилцааны дэг, тууштай байдал сулрах." },
    5: { title: "Уян хатан байдлын цоорхой", risk: "Өөрчлөлтөд хатуу хандаж мөргөлдөөн үүсгэх." },
    6: { title: "Үнэ цэнийн цоорхой", risk: "Өөрийгөө голж, хэт хамааралтай болох." },
    7: { title: "Итгэлцлийн цоорхой", risk: "Сэжиг нэмэгдэж, сэтгэлийн зай холдох." },
    8: { title: "Хил хязгаарын цоорхой", risk: "Үгүй гэж хэлж чадахгүй, ашиглуулах эрсдэлтэй." },
    9: { title: "Өнгөрснөө тавих цоорхой", risk: "Шарх гомдлоо удаан тээж шинэ харилцаа хаах." }
  },
  PERSONAL_YEAR_MAP: {
    1: { title: "Шинэ эхлэл" },
    2: { title: "Харилцан ойлголцол" },
    3: { title: "Нээлттэй илэрхийлэл" },
    4: { title: "Суурь ба тогтвортой байдал" },
    5: { title: "Өөрчлөлт ба хөдөлгөөн" },
    6: { title: "Хариуцлага ба халамж" },
    7: { title: "Дотоод цэвэрлэгээ" },
    8: { title: "Үр дүн ба статустай учрал" },
    9: { title: "Төгсгөл ба салалт/цэвэрлэгээ" }
  },
  RISK_DAY_GROUPS: {
    control: [8, 17, 26],
    unstable: [5, 14, 23],
    cold: [7, 16, 25]
  },
  LUCKY_ITEM_MAP: {
    1: { color: "Алтлаг, Улаан", stone: "Рубин" },
    2: { color: "Мөнгөлөг цагаан", stone: "Сарны чулуу" },
    3: { color: "Шар, Цайвар ягаан", stone: "Цитрин" },
    4: { color: "Хөх, Саарал", stone: "Лазурит" },
    5: { color: "Ногоон, Оюу", stone: "Маргад" },
    6: { color: "Ягаан, Цагаан", stone: "Сарны чулуу" },
    7: { color: "Гүн хөх, Нил ягаан", stone: "Аметист" },
    8: { color: "Хар, Хар хөх", stone: "Обсидиан" },
    9: { color: "Нил ягаан", stone: "Аметист" },
    11: { color: "Мөнгөлөг цагаан, Усан цэнхэр, Нил ягаан", stone: "Сарны чулуу, Аметист" },
    22: { color: "Ногоон, Алтлаг", stone: "Хаш" },
    33: { color: "Цэнхэр, Цагаан", stone: "Ларимар" }
  },
  LUCKY_NUMBERS_MAP: {
    1: [1, 10, 19, 28], 2: [2, 11, 20, 29], 3: [3, 12, 21, 30],
    4: [4, 13, 22, 31], 5: [5, 14, 23], 6: [6, 15, 24],
    7: [7, 16, 25], 8: [8, 17, 26], 9: [9, 18, 27],
    11: [1, 10, 11, 19, 20, 28, 29], 22: [2, 4, 8, 11, 20], 33: [3, 6, 9, 12, 21]
  }
};

function sumDigits(n) {
  return String(Math.abs(Number(n))).split("").reduce((a, b) => a + Number(b), 0);
}

function reduceNumber(n, keepMaster) {
  let v = Number(n);
  while (v > 9) {
    if (keepMaster && (v === 11 || v === 22 || v === 33)) break;
    v = sumDigits(v);
  }
  return v;
}

function calculateLifePath(y, m, d) {
  const yR = reduceNumber(y, true);
  const mR = reduceNumber(m, true);
  const dR = reduceNumber(d, true);
  const totalRaw = yR + mR + dR;
  const total = reduceNumber(totalRaw, true);
  return {
    number: total,
    rawSums: [dR, totalRaw],
    calculation: `(${y} -> ${yR}) + (${m} -> ${mR}) + (${d} -> ${dR}) = ${totalRaw} -> ${total}`
  };
}

function calculatePythagorasMatrix(y, m, d) {
  const dateStr = `${y}${String(m).padStart(2, "0")}${String(d).padStart(2, "0")}`;
  const dateDigits = dateStr.split("").map(Number);
  const n1 = dateDigits.reduce((a, b) => a + b, 0);
  const n2 = sumDigits(n1);
  const firstDigit = Number(String(d)[0]);
  const n3 = n1 - (2 * firstDigit);
  const n4 = sumDigits(Math.abs(n3));

  const allNums = [...dateDigits];
  String(n1).split("").forEach(c => allNums.push(Number(c)));
  String(n2).split("").forEach(c => allNums.push(Number(c)));
  String(Math.abs(n3)).split("").forEach(c => allNums.push(Number(c)));
  String(n4).split("").forEach(c => allNums.push(Number(c)));

  const counts = {};
  for (let i = 1; i <= 9; i++) counts[i] = 0;
  allNums.forEach(n => {
    if (counts[n] !== undefined) counts[n]++;
  });

  return { counts, rawN3: n3, rawN4: n4 };
}

function calculateHiddenNumbers(y, m, d, n3, n4) {
  const dateStr = `${y}${String(m).padStart(2, "0")}${String(d).padStart(2, "0")}`;
  const counts = {};
  const check = (n) => {
    String(Math.abs(n)).split("").forEach(c => {
      if (!dateStr.includes(c) && c !== "0") counts[c] = (counts[c] || 0) + 1;
    });
  };
  check(n3);
  check(n4);
  const list = Object.keys(counts).map(n => ({ number: Number(n), count: counts[n] }));
  return { list, text: list.map(i => `${i.number} (x${i.count})`).join(", ") };
}

function checkKarmicDebt(day, rawSums) {
  const pool = [Number(day)].concat((rawSums || []).map(Number));
  const debts = [13, 14, 16, 19].filter(n => pool.indexOf(n) !== -1);
  return debts.length ? debts.join(", ") : "Илрээгүй";
}

function calculatePersonalYear(month, day, year) {
  const md = reduceNumber(Number(month) + Number(day), false);
  const yy = reduceNumber(Number(year), false);
  return reduceNumber(md + yy, false);
}

function calculatePersonalYearForecast(month, day, startYear, span) {
  const years = [];
  const out = [];
  for (let i = 0; i < span; i++) {
    const y = Number(startYear) + i;
    const n = calculatePersonalYear(month, day, y);
    const title = (NUMEROLOGY.PERSONAL_YEAR_MAP[n] || { title: "Мөчлөг" }).title;
    years.push({ year: y, number: n, title: title });
    out.push(`${y} он -> Хувийн жил ${n} (${title})`);
  }
  return { years: years, text: out.join("; ") };
}

function buildCompatibilityNumbers(lifePathNumber, missingNumbers) {
  const missSet = {};
  (missingNumbers || []).forEach(item => {
    const n = Number(item && item.number);
    if (n >= 1 && n <= 9) missSet[n] = true;
  });

  const picks = [];
  if (missSet[6]) picks.push(6, 15, 24);
  if (missSet[8]) picks.push(8, 17, 26);
  if (lifePathNumber === 11 || lifePathNumber === 2) picks.push(2, 11, 20, 29);
  if (!picks.length) picks.push(2, 6, 11, 20, 24, 29);

  const seen = {};
  return picks.filter(x => {
    if (seen[x]) return false;
    seen[x] = true;
    return true;
  }).join(", ");
}

function buildRiskProfiles() {
  const c = NUMEROLOGY.RISK_DAY_GROUPS.control.join(", ");
  const u = NUMEROLOGY.RISK_DAY_GROUPS.unstable.join(", ");
  const cold = NUMEROLOGY.RISK_DAY_GROUPS.cold.join(", ");
  return `Хэт хяналттай энерги: ${c}; Тогтворгүй энерги: ${u}; Сэтгэлээ хаадаг энерги: ${cold}`;
}

function buildMatrixSummary(counts) {
  const chunks = [];
  for (let i = 1; i <= 9; i++) {
    chunks.push(`${i}:${counts[i]}`);
  }
  return chunks.join(" | ");
}

function calculateAge(y, m, d) {
  const now = new Date();
  let age = now.getFullYear() - y;
  const currentMonth = now.getMonth() + 1;
  const currentDay = now.getDate();
  if (currentMonth < m || (currentMonth === m && currentDay < d)) age--;
  return age;
}

// ==========================================
// 3. AI PARSING ENGINE (Normalization)
// ==========================================
function parseAndCalculateProfile(rawInput, apiKey) {
  const normalized = normalizeInputWithAI(rawInput, apiKey);
  const [year, month, day] = normalized.date.split(".").map(Number);

  const lifePath = calculateLifePath(year, month, day);
  const lifePathInfo = NUMEROLOGY.LIFE_PATH_MAP[lifePath.number] || NUMEROLOGY.LIFE_PATH_MAP[1];
  const dayNumber = reduceNumber(day, false);
  const matrix = calculatePythagorasMatrix(year, month, day);
  const hiddenNumbers = calculateHiddenNumbers(year, month, day, matrix.rawN3, matrix.rawN4);
  const karmicDebt = checkKarmicDebt(day, lifePath.rawSums);

  const age = calculateAge(year, month, day);
  const forecast = calculatePersonalYearForecast(month, day, 2026, 3);

  const excessNumbers = [];
  for (let n = 1; n <= 9; n++) {
    if (matrix.counts[n] >= 3) {
      excessNumbers.push({
        number: n,
        count: matrix.counts[n],
        title: NUMEROLOGY.MATRIX_EXCESS_MAP[n].title,
        desc: NUMEROLOGY.MATRIX_EXCESS_MAP[n].desc
      });
    }
  }

  const missingNumbers = [];
  for (let n = 1; n <= 9; n++) {
    if (matrix.counts[n] === 0) {
      missingNumbers.push({
        number: n,
        title: NUMEROLOGY.MISSING_NUMBER_MAP[n].title,
        risk: NUMEROLOGY.MISSING_NUMBER_MAP[n].risk
      });
    }
  }

  const compatibleNumbers = buildCompatibilityNumbers(lifePath.number, missingNumbers);
  const riskProfiles = buildRiskProfiles();

  const lucky = NUMEROLOGY.LUCKY_ITEM_MAP[lifePath.number] || NUMEROLOGY.LUCKY_ITEM_MAP[11];
  const luckyNumbers = NUMEROLOGY.LUCKY_NUMBERS_MAP[lifePath.number] || NUMEROLOGY.LUCKY_NUMBERS_MAP[11];

  return {
    name: normalized.name,
    gender: normalized.gender,
    dob: normalized.date,
    age: age,
    parsingUsage: normalized.usage,

    lifePath: lifePath.number,
    lifePathName: lifePathInfo.name,
    lifePathDesc: lifePathInfo.desc,
    lifePathCalcString: lifePath.calculation,
    dayNumber: dayNumber,

    matrixCounts: matrix.counts,
    matrixSummary: buildMatrixSummary(matrix.counts),
    excessNumbers: excessNumbers,
    missingNumbers: missingNumbers,
    hiddenNumbers: hiddenNumbers,
    karmicDebt: karmicDebt,

    forecastText: forecast.text,
    forecastYears: forecast.years,

    compatibleNumbers: compatibleNumbers,
    riskProfiles: riskProfiles,

    luckyColor: lucky.color,
    luckyStone: lucky.stone,
    luckyNumbers: luckyNumbers.join(", ")
  };
}

function normalizeInputWithAI(raw, key) {
  const prompt = `
  TASK: Convert the input into strict JSON.
  INPUT: "${raw}"

  RULES:
  - Date must be YYYY.MM.DD
  - Gender must be "Эрэгтэй" or "Эмэгтэй" (Infer from names/words like 'эр', 'эм', 'хүү', 'охин'. Default "Эмэгтэй").
  - Name optional, default "Эрхэм"
  - Return JSON only, no markdown.

  JSON FORMAT:
  {"date":"YYYY.MM.DD","gender":"Эрэгтэй","name":"Эрхэм"}
  `;
  try {
    const result = callGeminiAPI(prompt, key, 0.1, true);
    const aiData = JSON.parse(result.text.trim());
    if (!aiData.date) throw new Error("No date in JSON");

    let gender = "Эмэгтэй";
    let rawG = String(aiData.gender || "").toLowerCase();
    if (rawG.includes("эр") || rawG === "male" || rawG === "man") gender = "Эрэгтэй";

    let name = aiData.name || "Эрхэм";
    if (name.length > 30) name = "Эрхэм";

    return {
        date: aiData.date.replace(/[\s\-\/]/g, "."),
        gender: gender,
        name: name,
        usage: result.usage
    };
  } catch (e) {
    const dates = raw.match(/\d{4}[\.\-\s\/]\d{1,2}[\.\-\s\/]\d{1,2}/g) || ["1990.01.01"];
    let fallbackGender = "Эмэгтэй";
    if (raw.toLowerCase().includes("эр")) fallbackGender = "Эрэгтэй";

    return {
      date: dates[0].replace(/[\s\-\/]/g, ".") || "1990.01.01",
      gender: fallbackGender,
      name: "Эрхэм",
      usage: 0
    };
  }
}

// ==================================================================================
// 4. GENERATION PROMPTS & REFERENCES (REFINED PERSONA)
// ==================================================================================
CONFIG.REFERENCES = {
    PART_1: `1-р хэсэг. ХАЙРЫН АРХЕТИП: Сэтгэл Зүйн Зураг Төөрөг
🔮 Таны төрсөн он сар өдрийн нийлбэр тооцооллоор [Мастер тоо / Амьдралын зам] гарч байгаа нь таныг хайр дурлалын ертөнцөд [Архетипийн нэр] гэсэн маш онцгой мэдрэмтгий дүр төрхөөр тодорхойлж байна. Энэ нь таныг жирийн нэг хүн биш, харин эсрэг хүнийхээ дотоод ертөнц рүү өнгийж хардаг, сэтгэлийн нандин холбоог бүхнээс илүүд үздэг хүн гэдгийг илтгэдэг юм.
👑 Нэгдүгээрт, таны тоон өгөгдөлд [Тоо] буюу [Цифр]-ийн цифр [Тоо] удаа давтагдан байрлаж байгаа нь танд төрөлхийн маш хүчтэй [Зан чанар]-г өгдөг. Энэ нь таныг бие даасан болгодог сайн талтай ч, харилцаанд энэ хүч хэтэрвэл эсрэг хүнээ өөрийн мэдэлгүй захирах эсвэл шүүмжилж голдог сүүдэр талыг үүсгэдэг тул та энэ хүчээ тэнцвэржүүлэх шаардлагатай.
🗣️ Хоёрдугаарт, таны өдрийн тоо [Тоо] байгаа нь таны Хайрын хэлийг [Хайрын хэл] гэж тодорхойлж байна. Танд зүгээр нэг үнэтэй бэлэг өгдөг хүнээс илүү тантай орой бүр цаг гарган ярилцдаг, таныг байнга урамшуулдаг хүн хамгийн их таалагддаг бөгөөд таныг чин сэтгэлээсээ сонсдог хүн л таны зүрхийг эзэмшиж чадна.`,

    PART_2: `2-р хэсэг. ГАНЦ БИЕ БАЙДЛЫН ОНОШ: Анхааруулах Дохионууд
🚧 Таны төрсөн он сар өдрийн тооцооллыг нарийвчлан шалгаж үзэхэд [Үйлийн үрийн тоо] илэрч байна (эсвэл илрээгүй нь давуу тал юм). Энэ нь таныг хайр сэтгэлийн харилцаанд [Сөрөг/Эерэг] ачаа тээш тээж явааг харуулж байгаа тул одоогийн нөхцөл байдал таны [Сонголт/Өнгөрсөн үйлдэл]-ээс шууд хамаарч байна.
🕳️ Таны тоон мэдээлэлд [Тоо] дутуу байгаа нь таны хайр дурлалын харилцаанд [Сэтгэл зүйн гацаа] үүсгэж байна. Энэ нь таныг хайр дурлалд орохоороо өөрийгөө бүрэн мартаж хамааралтай болох, эсвэл өөрийн орон зайг хамгаалж чадахгүй байх эрсдэлд хүргэдэг.
🔒 Сэтгэл зүйн хувьд та дотоод сэтгэлдээ ирээдүйн ханиа маш өндөр шалгуураар төсөөлдөг, эсвэл өнгөрсөн харилцааны дурсамжаас бүрэн салж чадахгүй зууралдсаар байх магадлалтай. Шинээр танилцсан хүн бүрийг ухамсаргүйгээр өнгөрсөнтэйгөө харьцуулах нь шинэ харилцааг хаах гол шалтгаан болдог.
⚠️ Анхааруулах дохионы хувьд [Эрсдэлтэй өдрүүд] өдөр төрсөн буюу [Эрсдэлийн зан чанар] энергитэй хүн таны дотоод амар амгаланг эвдэх магадлалтай тул та ийм харилцаанаас сэтгэл зүйгээ хамгаалах хэрэгтэй.`,

    PART_3: `3-р хэсэг. ИРЭЭДҮЙН ХАНИЙН ДҮР ТӨРХ: Таны Төгс Зохицол
💎 Таны заяаны ханийн дүр төрхийг тооцоолж үзэхэд танд тохирох төгс зохицол нь таны дотоод шуургыг намжаадаг, сэтгэл зүйн хувьд маш тогтвортой "Дөлгөөн Халамжлагч" дүр төрхтэй хүн байх болно. Тэр хүн олны дунд хэт чанга дуугарч өөрийгөө дөвийлгөдөггүй, харин цаанаасаа л нэг тийм намуухан, уужуу тайван, бусдыг хүндэлдэг соёлтой зан чанартай байх бөгөөд түүний хажууд байхад танд өөрийнхөө тэр их хүчийг гаргах шаардлагагүй, яг л аюулгүй бүсэд байгаа мэт тайван мэдрэмж төрнө.
💼 Мэргэжил болон нийгмийн байдлын хувьд тэр хүн мөнгөний хойноос улайрч явдаг хүн гэхээсээ илүүтэй бусдын төлөө үйлчилдэг, оюуны хөдөлмөр эрхэлдэг эсвэл бүтээлч салбарын хүн байх магадлал маш өндөр байна.
🔢 Таны энергитэй хамгийн төгс зохицох алтан тоонууд буюу ирээдүйн ханийн тань төрсөн өдөр нь сарын [Зохицох тоонууд]-ний өдрүүд байх магадлалтай. Эдгээр хүмүүс таны дутууг нөхөж, танд байхгүй тэрхүү гэр бүлсэг, тогтвортой байдлыг бэлэглэж чадна.
✨ Гадаад төрхийн хувьд тэр хүн тийм ч хурц ширүүн харцтай биш, зөөлөн дулаан харцтай, цэвэрч нямбай хувцасладаг хүн байх бөгөөд та түүнийг анх харахад л маш дотно, танил мэдрэмж төрөх болно.`,

    PART_4: `4-р хэсэг. ХАЙР ДУРЛАЛЫН 3 ЖИЛИЙН ПРОГНОЗ: Цаг Хугацааны Зураглал
⏳ Таны [Жил] оны хувийн энергийн мөчлөгийг тооцоолж үзэхэд та энэ жил [Утга] гэсэн маш хүчирхэг мөчлөгт орж байна. Энэ нь таны хувьд хайр дурлалын түүхэндээ цоо шинэ хуудас нээх онцгой жил байх болно. Хэрэв та дотоод сэтгэлээ цэгцэлж чадсан бол яг энэ жил амжилттай яваа хүчирхэг хүнтэй учрах магадлал хамгийн өндөр байна.
🧹 Харин дараагийн жил буюу [Жил] он нь [Утга] мөчлөг тохиох бөгөөд энэ нь таны хайр дурлалын амьдралд их цэвэрлэгээ хийгдэх жил байх болно. Өмнөх бүх эргэлзээ, айдас болон хуучин дурсамжуудаасаа бүрмөсөн салж, зөвхөн өөрийн сонгосон тэр хүнтэйгээ ирээдүйгээ холбох сэтгэл зүйн бэлтгэлээ хангах үе юм.
💒 Эцэст нь [Жил] он бол таны хувьд цоо шинэ мөчлөгийн эхлэл буюу [Утга] тохиож байгаа нь албан ёсоор гэр бүл болох эсвэл хамтын амьдралаа эхлүүлэхэд хамгийн ивээлтэй жил байх болно.
📍 Таны мэдрэмтгий энерги нь хэт их чимээ шуугиантай газруудад хаагддаг тул та ирээдүйн ханиа баар цэнгээний газраас бус, харин өөрийнхөө сүнслэг болон оюунлаг мөн чанарт тохирсон орчин болох хувь хүний хөгжлийн сургалт, номын нээлт эсвэл ажил хэргийн уулзалтууд дээрээс олох магадлал өндөр байна.`,

    PART_5: `5-р хэсэг. АМЖИЛТЫН ТҮЛХҮҮР: Тэнцвэр Ба Хамгаалалт
🔑 Таны хайр дурлалын амьдралыг өөрчлөх хамгийн чухал стратеги бол буулт хийж сурах урлаг юм. Таны хүчирхэг өгөгдөл нь таныг бүх зүйлийг хяналтдаа байлгах зуршилтай болгосон байдаг тул та хайр дурлалын харилцаанд орохоороо өөрийн мэдэлгүй удирдах дүрд тоглож эхэлдэг. Тиймээс та зөөлөн, хүлээн авагч чанараа ил гаргаж сурах нь таныг жинхэнэ халамжтай хүнтэй учрахад туслах болно. Мөн хэт их бодох зуршлыг дарахын тулд тархиа амраах дасгалуудыг тогтмол хийх хэрэгтэй.
🧘‍♀️ Таны далд ухамсарт суусан "Би ганцаараа байх ёстой" эсвэл "Намайг хэн ч ойлгохгүй" гэх мэт сөрөг итгэл үнэмшлийг эвдэхийн тулд "Би өнгөрсөн бүх гомдлоо тавьж явууллаа, би одоо шинэ хайрыг хүлээж авахад бэлэн байна" гэж өөртөө хэлж хэвшээрэй.
🎨 Таны энергийг хамгаалж, зөн совинг тань улам хурцлахын тулд та [Өнгө] өнгийн хувцас эсвэл хэрэглэл ашиглах нь маш эерэг нөлөөтэй. Харин таны энергийг цэвэрлэж, хайр дурлалыг дуудах байгалийн чулуу бол [Чулуу] юм.
📜 Эцсийн дүгнэлтээр, та бол хорвоо ертөнцийг зүрхээрээ мэдэрдэг онцгой өгөгдөлтэй хүн юм. Таны ганц бие байгаа шалтгаан нь таны шалгуур өндөр, таны сэтгэл зүгээр нэг энгийн харилцааг биш гүн гүнзгий холбоог хайж байгаад оршино. Одоо та өөрийнхөө хүчийг зөөлөлж, ухаанаа амрааж чадвал таныг хувь заяаны гайхалтай бэлэг хүлээж байна.`
};

// ==========================================
// 5. AI GENERATION ENGINE (5-Part Sequential, Token-Safe)
// ==========================================
function generateSequentialReport(data, apiKey) {
  const header = `
💎 ХЭРЭГЛЭГЧИЙН ЗУРАГ ТӨӨРӨГ:
Нэр: ${data.name} | Хүйс: ${data.gender} | Нас: ${data.age} | Төрсөн огноо: ${data.dob}
Амьдралын зам: ${data.lifePath} (${data.lifePathName}) | Өдрийн тоо: ${data.dayNumber}
Матрицын бүтэц: ${data.matrixSummary} | Дутуу тоо: ${data.missingNumbers.map(m=>m.number).join(",") || "Байхгүй"}
Кармын өр: ${data.karmicDebt}
`;

  const SYSTEM_PROMPT = `
  ROLE: You are an elite, highly intuitive Numerologist & Psychotherapist. Tone: Direct, Grounded, Psychological, and Wise. Avoid excessive analogies and fairy-tale similes. Write as a mature, experienced counselor speaking directly to the user about real life.
  LANGUAGE: Proper Mongolian Cyrillic ONLY. Avoid foreign words where possible.

  >>> MASTER RULES (STRICTLY ENFORCED): <<<
  1. ZERO META-TALK: NEVER use phrases like "Энэ хэсэгт бид...", "Дүгнэж хэлэхэд...", "За ойлголоо". Start your analysis IMMEDIATELY.
  2. NO POETRY / FAIRY TALES: Do NOT write like a poem. Be realistic, deeply psychological, and analytical.
  3. MONGOLIAN GRAMMAR: Respect user's gender strictly ("эрэгтэй", "эмэгтэй"). If wording is uncertain, use neutral "хүн".
  4. STRICT EMOJI RULE: EVERY SINGLE PARAGRAPH (except the section title) MUST start with EXACTLY ONE emoji (e.g., 🔮, 👑, 🗣️). ZERO exceptions. Do NOT use any emojis in the middle or end of sentences.
  5. STRICT FORMATTING:
     - NO Markdown headers like (#, ##).
     - NO Markdown bold formatting like (**text**).
     - Return ONLY PLAIN TEXT separated by double line breaks.
  6. COMPLETENESS: NEVER cut off mid-sentence. Always finish your thoughts and provide a complete, well-rounded conclusion. Aim for around 450 words per part.
  `;

  const excessStr = data.excessNumbers.map(n => n.number).join(", ") || "Байхгүй";
  const missingStr = data.missingNumbers.map(n => n.number).join(", ") || "Байхгүй";

  // PART 1: LOVE DNA
  const prompt1 = `
  ${SYSTEM_PROMPT}
  TASK: Write PART 1 ONLY (Love Archetype & Matrix).

  DATA:
  - Амьдралын зам: ${data.lifePath} (${data.lifePathName})
  - Өдрийн тоо: ${data.dayNumber}
  - Матрицад хэт хүчтэй байгаа тоонууд: ${excessStr}

  INSTRUCTIONS: Begin exactly with "1-р хэсэг. ХАЙРЫН АРХЕТИП: Сэтгэл Зүйн Зураг Төөрөг" on its own line, then double line break, then start.
  First paragraph: Explain their Love Archetype based on their Life Path (${data.lifePath}). Second paragraph: Dive into the psychological traits of their Matrix Excess numbers (${excessStr}). Explain how this extreme energy affects their relationship style. Third paragraph: Explain their Love Language based on their Day Number (${data.dayNumber}). Final paragraph: A practical psychological conclusion.

  STYLE GUIDE REFERENCE (Model your structure, depth, and tone exactly after this):
  ${CONFIG.REFERENCES.PART_1}
  `;
  const r1 = callGeminiAPI(prompt1, apiKey, CONFIG.TEMPERATURE);

  // PART 2: DIAGNOSIS
  const prompt2 = `
  ${SYSTEM_PROMPT}
  TASK: Write PART 2 ONLY (Single Status Diagnosis & Danger Signals).

  DATA:
  - Кармын өр: ${data.karmicDebt}
  - Матрицын дутуу тоонууд: ${missingStr}
  - Аюулын бүс (Эрсдэлийн энерги): ${data.riskProfiles}

  INSTRUCTIONS: Begin exactly with "2-р хэсэг. ГАНЦ БИЕ БАЙДЛЫН ОНОШ: Анхааруулах Дохионууд" on its own line, then double line break, then start.
  First paragraph: Address their Karmic Debt (${data.karmicDebt}). Is it a burden or a clean slate? Second paragraph: Discuss their Missing Numbers (${missingStr}) and how it causes a psychological block or lack of boundaries in love. Third paragraph: Discuss their mindset block (idealism vs past memories). Fourth paragraph: Warn them about the Danger Zones (${data.riskProfiles})—what kind of people disrupt their peace.

  STYLE GUIDE REFERENCE (Model your structure, depth, and tone exactly after this):
  ${CONFIG.REFERENCES.PART_2}
  `;
  const r2 = callGeminiAPI(prompt2, apiKey, CONFIG.TEMPERATURE);

  // PART 3: PARTNER AVATAR
  const prompt3 = `
  ${SYSTEM_PROMPT}
  TASK: Write PART 3 ONLY (Future Partner Avatar).

  DATA:
  - Төгс зохицох өдрийн тоонууд: ${data.compatibleNumbers}
  - Матрицын дутуу тоог нөхөх хэрэгцээ: ${missingStr}

  INSTRUCTIONS: Begin exactly with "3-р хэсэг. ИРЭЭДҮЙН ХАНИЙН ДҮР ТӨРХ: Таны Төгс Зохицол" on its own line, then double line break, then start.
  First paragraph: Describe the personality archetype of their perfect match. Someone who grounds them and heals their specific missing numbers (${missingStr}). Second paragraph: Discuss the likely profession, social status, and mindset of this person. Third paragraph: List the specific birth dates (${data.compatibleNumbers}) that match them best and why. Fourth paragraph: Describe their physical vibe and the initial feeling of meeting them.

  STYLE GUIDE REFERENCE (Model your structure, depth, and tone exactly after this):
  ${CONFIG.REFERENCES.PART_3}
  `;
  const r3 = callGeminiAPI(prompt3, apiKey, CONFIG.TEMPERATURE);

  // PART 4: 3-YEAR FORECAST
  const prompt4 = `
  ${SYSTEM_PROMPT}
  TASK: Write PART 4 ONLY (3-Year Forecast).

  DATA:
  - 3 жилийн хувийн мөчлөг: ${data.forecastText}

  INSTRUCTIONS: Begin exactly with "4-р хэсэг. ХАЙР ДУРЛАЛЫН 3 ЖИЛИЙН ПРОГНОЗ: Цаг Хугацааны Зураглал" on its own line, then double line break, then start.
  Address each of the 3 years individually. Explain what the specific "Хувийн жил" (Personal Year) means for their love life. E.g., is it a year of karma, a year of clearing the past, or a year of marriage? Final paragraph: Suggest the best physical environments or situations where they are most likely to meet their soulmate based on their refined energy.

  STYLE GUIDE REFERENCE (Model your structure, depth, and tone exactly after this):
  ${CONFIG.REFERENCES.PART_4}
  `;
  const r4 = callGeminiAPI(prompt4, apiKey, CONFIG.TEMPERATURE);

  // PART 5: SUCCESS KEY
  const prompt5 = `
  ${SYSTEM_PROMPT}
  TASK: Write PART 5 ONLY (Success Strategy & Conclusion).

  DATA:
  - Амжилтын өнгө: ${data.luckyColor}
  - Азын чулуу: ${data.luckyStone}

  INSTRUCTIONS: Begin exactly with "5-р хэсэг. АМЖИЛТЫН ТҮЛХҮҮР: Тэнцвэр Ба Хамгаалалт" on its own line, then double line break, then start.
  First paragraph: The psychological strategy to succeed (e.g., surrendering control, stopping overthinking). Second paragraph: A morning affirmation/mantra to break their subconscious blocks. Third paragraph: Their Lucky Colors (${data.luckyColor}) and Lucky Stones (${data.luckyStone}) to protect their energy. Final paragraph: A strong, empowering psychological conclusion summarizing their unique Numerology blueprint.

  STYLE GUIDE REFERENCE (Model your structure, depth, and tone exactly after this):
  ${CONFIG.REFERENCES.PART_5}
  `;
  const r5 = callGeminiAPI(prompt5, apiKey, CONFIG.TEMPERATURE);

  return {
    text: header.trim() + "\n\n" + r1.text.trim() + "\n\n" + r2.text.trim() + "\n\n" + r3.text.trim() + "\n\n" + r4.text.trim() + "\n\n" + r5.text.trim(),
    usage: (r1.usage||0) + (r2.usage||0) + (r3.usage||0) + (r4.usage||0) + (r5.usage||0)
  };
}

// ==========================================
// 6. API CALLER (GEMINI) - Robust Retry Engine
// ==========================================
function callGeminiAPI(prompt, apiKey, temp, requireJson = false) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.GEMINI_MODEL}:generateContent?key=${apiKey}`;

  let payload = {
    "contents": [{ "parts": [{ "text": prompt }] }],
    "generationConfig": { "temperature": temp, "maxOutputTokens": 8192 },
    "safetySettings": [
        { "category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_ONLY_HIGH" },
        { "category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_ONLY_HIGH" },
        { "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_ONLY_HIGH" },
        { "category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_ONLY_HIGH" }
    ]
  };

  if (requireJson) payload.generationConfig.responseMimeType = "application/json";

  const maxAttempts = 3;
  let lastErrorMsg = "";

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const res = UrlFetchApp.fetch(url, { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true });

    if (res.getResponseCode() === 429 || res.getResponseCode() >= 500) {
        lastErrorMsg = res.getContentText();
        Utilities.sleep(attempt * 2000);
        continue;
    }

    try {
        const json = JSON.parse(res.getContentText());
        if (json.candidates && json.candidates[0].content) {
            return {
                text: json.candidates[0].content.parts[0].text,
                usage: json.usageMetadata ? json.usageMetadata.totalTokenCount : 0
            };
        }
        lastErrorMsg = res.getContentText();
    } catch(e) {
        lastErrorMsg = e.toString() + " | " + res.getContentText();
        Utilities.sleep(attempt * 2000);
    }
  }

  throw new Error(`Gemini API Error after ${maxAttempts} attempts: ${lastErrorMsg}`);
}

// ==========================================
// 7. SAFE PDF DELIVERY ENGINE & UCHAT
// ==========================================
function createPdfSafely(name, content, templateId, folderId) {
  const targetFolder = DriveApp.getFolderById(folderId);
  const copyFile = DriveApp.getFileById(templateId).makeCopy(`${name} - ${CONFIG.PRODUCT_NAME}`, targetFolder);
  const copyId = copyFile.getId();

  const doc = DocumentApp.openById(copyId);
  const body = doc.getBody();

  body.replaceText("(?i){{name}}", name);

  let cleanText = content.replace(/\*/g, "");
  cleanText = cleanText.replace(/^#+\s/gm, "");

  const paragraphs = cleanText.split(/\n+/);
  for (let i = 0; i < paragraphs.length; i++) {
    let pText = paragraphs[i].trim();
    if (pText.length > 0) {
      const firstCharMatch = pText.match(/^[\u{1F300}-\u{1F6FF}\u{1F900}-\u{1F9FF}\u{2600}-\u{26FF}\u{2700}-\u{27BF}\u{1F1E6}-\u{1F1FF}\u{1F200}-\u{1F2FF}]/u);
      let firstEmoji = firstCharMatch ? firstCharMatch[0] : "";
      let noEmojiText = pText.replace(/[\u{1F300}-\u{1F6FF}\u{1F900}-\u{1F9FF}\u{2600}-\u{26FF}\u{2700}-\u{27BF}\u{1F1E6}-\u{1F1FF}\u{1F200}-\u{1F2FF}]/gu, "");
      paragraphs[i] = (firstEmoji + " " + noEmojiText).trim();
    }
  }

  // Rule 6: Locate and replace the EXACT placeholder safely
  let insertionIndex = -1;
  let textAttributes = {};

  const searchResult = body.findText("(?i){{\\s*report\\s*}}");

  if (searchResult) {
    const element = searchResult.getElement();
    const textElement = element.asText();

    // Extract exact font formatting to inherit
    textAttributes = textElement.getAttributes();

    const paragraphToReplace = element.getParent();
    insertionIndex = body.getChildIndex(paragraphToReplace);
    paragraphToReplace.removeFromParent();
  } else {
    insertionIndex = body.getNumChildren() - 1;
  }

  for (let i = paragraphs.length - 1; i >= 0; i--) {
    let pText = paragraphs[i].trim();
    if (pText.length > 0) {
        let p = body.insertParagraph(insertionIndex, pText);

        // Inherit user's custom font and size from {{report}} placeholder
        let pTextElement = p.editAsText();

        // Clean inherited attributes to only keep FontFamily and FontSize
        if (textAttributes[DocumentApp.Attribute.FONT_FAMILY]) {
            pTextElement.setFontFamily(textAttributes[DocumentApp.Attribute.FONT_FAMILY]);
        }
        if (textAttributes[DocumentApp.Attribute.FONT_SIZE]) {
            pTextElement.setFontSize(textAttributes[DocumentApp.Attribute.FONT_SIZE]);
        }

        // Alignment
        if (pText.length > 50) {
            p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
        }

        if (/^\d-р хэсэг/.test(pText)) {
            // Only bold and slightly increase size, preserving user's font (e.g. Oswald)
            pTextElement.setBold(true);
            if (textAttributes[DocumentApp.Attribute.FONT_SIZE]) {
                pTextElement.setFontSize(textAttributes[DocumentApp.Attribute.FONT_SIZE] + 2); // e.g. 13 -> 15
            } else {
                pTextElement.setFontSize(14);
            }
            p.setSpacingBefore(20);
            p.setSpacingAfter(10);
        } else {
            p.setLineSpacing(1.5);
            p.setSpacingAfter(10);
        }
    }
  }

  doc.saveAndClose();

  const pdfBlob = copyFile.getAs('application/pdf');
  const pdfFile = targetFolder.createFile(pdfBlob);
  pdfFile.setName(`${name} - Тайлан.pdf`);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  copyFile.setTrashed(true);
  return pdfFile.getUrl();
}

function sendUChatProven(userNs, pdfUrl, name, token) {
  if (!token) throw new Error("DELIVERY: UChat token байхгүй.");
  if (!userNs) throw new Error("DELIVERY: user_ns хоосон.");

  const msg = CONFIG.UCHAT.DELIVERY_MESSAGE.replace(/\{\{NAME\}\}/g, name);
  const btn = CONFIG.UCHAT.DELIVERY_BTN_TEXT;

  const payload = {
    user_ns: String(userNs).trim(),
    data: {
      version: "v1",
      content: { messages: [ { type: "text", text: msg, buttons: [ { type: "url", caption: btn, url: pdfUrl } ] } ] }
    }
  };

  const res = UrlFetchApp.fetch(CONFIG.UCHAT.ENDPOINT, {
    method: "post",
    headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });

  const status = res.getResponseCode();
  const body = res.getContentText();

  if (status < 200 || status >= 300) throw new Error("DELIVERY HTTP " + status + ": " + body.substring(0, 200));
  const json = JSON.parse(body);
  if (json.status !== "ok" && json.success !== true) throw new Error("DELIVERY API failed: " + JSON.stringify(json));
}

// --- КОДЫН ТӨГСГӨЛ ---
