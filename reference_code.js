// --- КОДЫН ЭХЛЭЛ ---

/****************************************************************************************
 * PRODUCT: DIGITAL ASTROLOGY REPORT GENERATOR (MONEY & KARMA EDITION)
 * ENGINE: ROBUST V26.1 (Smart Doc Insertion, Token Limit Safe, UChat, Regex Cleaner)
 ****************************************************************************************/

const CONFIG = {
  VERSION: "v26.1-Astrology-Pro",
  PRODUCT_NAME: "Хувь Заяаны Алтан Түлхүүр: Эд Баялгийн Зурхай",
  SHEET_NAME: "Sheet1",
  BATCH_SIZE: 3, 
  GEMINI_MODEL: "gemini-2.5-flash", 
  TEMPERATURE: 0.7, 

  COLUMNS: {
    NAME: 0, ID: 1, INPUT: 2, PDF: 3, STATUS: 4, 
    TOKEN: 5, DEBUG: 6, DATE: 7, VER: 8, ERROR: 9
  },

  UCHAT: {
    ENDPOINT: "https://www.uchat.com.au/api/subscriber/send-content",
    DELIVERY_MESSAGE: `Сайн байна уу, {{NAME}}? 🔮\n\nТаны "Хувь Заяаны Алтан Түлхүүр: Эд Баялгийн Зурхай" хувийн тайлан бэлэн боллоо.\n\nДоорх товч дээр дарж татаж авна уу. 👇`,
    DELIVERY_BTN_TEXT: "📥 Зурхай татах"
  },

  LIFE_PATH_MAP: {
    1: "Удирдагч (The Leader)", 2: "Зуучлагч (The Diplomat)", 3: "Бүтээгч (The Creator)", 4: "Гүйцэтгэгч (The Builder)",
    5: "Аялагч (The Adventurer)", 6: "Халамжлагч (The Nurturer)", 7: "Шинжээч (The Seeker)", 8: "Эрх мэдэлтэн (The Executive)",
    9: "Нигүүлсэгч (The Humanitarian)", 11: "Зөн Билэгтэн (The Visionary) - Их Хүчит Тоо", 
    22: "Их Бүтээн байгуулагч (The Master Builder) - Их Хүчит Тоо", 33: "Их Багш (The Master Teacher) - Их Хүчит Тоо"
  },
  ELEMENT_MAP: { "Гал": "Дүрэлзсэн Гал", "Төмөр": "Ган Төмөр", "Ус": "Ундрах Ус", "Мод": "Ургах Мод", "Шороо": "Шимт Шороо" },
  ELEMENTS_BY_LAST_DIGIT: { 0: "Төмөр", 1: "Төмөр", 2: "Ус", 3: "Ус", 4: "Мод", 5: "Мод", 6: "Гал", 7: "Гал", 8: "Шороо", 9: "Шороо" },
  
  TSAGAAN_SAR: {
    1940: "02-08", 1941: "01-27", 1942: "02-15", 1943: "02-05", 1944: "02-24", 1945: "02-13", 1946: "02-02", 1947: "01-22", 1948: "02-10", 1949: "01-29",
    1950: "02-17", 1951: "02-06", 1952: "01-27", 1953: "02-14", 1954: "02-03", 1955: "02-24", 1956: "02-12", 1957: "01-31", 1958: "02-18", 1959: "02-08",
    1960: "02-27", 1961: "02-15", 1962: "02-05", 1963: "02-25", 1964: "02-13", 1965: "02-02", 1966: "02-21", 1967: "02-09", 1968: "01-30", 1969: "02-17",
    1970: "02-06", 1971: "02-27", 1972: "02-15", 1973: "02-06", 1974: "02-23", 1975: "02-11", 1976: "01-31", 1977: "02-18", 1978: "02-07", 1979: "02-28",
    1980: "02-16", 1981: "02-05", 1982: "02-24", 1983: "02-13", 1984: "02-02", 1985: "02-20", 1986: "02-09", 1987: "01-29", 1988: "02-17", 1989: "02-06",
    1990: "02-27", 1991: "02-15", 1992: "02-04", 1993: "02-23", 1994: "02-10", 1995: "01-31", 1996: "02-19", 1997: "02-07", 1998: "02-28", 1999: "02-16",
    2000: "02-05", 2001: "02-24", 2002: "02-12", 2003: "02-01", 2004: "02-22", 2005: "02-09", 2006: "01-29", 2007: "02-18", 2008: "02-07", 2009: "02-25",
    2010: "02-14", 2011: "02-03", 2012: "02-22", 2013: "02-11", 2014: "01-31", 2015: "02-19", 2016: "02-09", 2017: "02-27", 2018: "02-16", 2019: "02-05",
    2020: "02-24", 2021: "02-12", 2022: "02-02", 2023: "02-21", 2024: "02-10", 2025: "02-28" 
  },

  MISSING_NUM_MAP: {
    1: "Өөрийн үзэл бодлыг чөлөөтэй илэрхийлж, бусдад 'Үгүй' гэж хэлж сурах хэрэгтэй. Бие даасан байдал, манлайлал дутагдаж байна.",
    2: "Бусдын оронд өөрийгөө тавьж, мэдрэмжтэй харилцах, багаар ажиллах чадварт суралцах шаардлагатай.",
    3: "Амьдралыг хэт хуурай, логикоор харахаа боль. Төсөөлөн бодох чадвар, бүтээлч сэтгэлгээгээ хөгжүүл.",
    4: "Ажил амьдралдаа эмх цэгц, нарийн төлөвлөгөөг эрхэмлэ. Аливааг заавал дуустал нь хийх тэвчээрт суралц.",
    5: "Өөрчлөлтөөс бүү ай. Шинэ зүйл туршиж, хуучин хэв маягаасаа салж, уян хатан байх нь танд боломж авчирна.",
    6: "Гэр бүл, дотны хүмүүстээ цаг гарга. Хариуцлагаас зугтах биш, түүнийг үүрч, бусдыг халамжилж сур.",
    7: "Дотоод сэтгэлээ сонс. Материаллаг зүйлээс гадна оюун санааны хөгжилдөө анхаарч, аливааг гүнзгий судал.",
    8: "Санхүүгийн сахилга бат, мөнгөний менежментэд зориуд суралц. Мөнгөнд хэт хэнэггүй хандаж, боломжийг алдахаас сэргийл.",
    9: "Бусдыг энэрч, өгөөмөр бай. Зөвхөн өөрийгөө бодох нь таны амжилтыг хязгаарлана. Нийгэмд тустай үйл хий."
  },

  PERSONAL_YEAR_MAP: {
    1: "Шинэ Эхлэл ба Үр Тарих. Энэ жил бол таны амьдралын шинэ мөчлөгийн эхлэл. Хуучин зүйлсээ ард үлдээж, цоо шинэ төсөл, ажил эхлүүлэхэд хамгийн таатай цаг. Өөртөө итгэлтэйгээр хөрөнгө оруулалт хийх, шинэ бизнес эхлүүлэх, тушаал дэвших хүсэлт гаргахад тохиромжтой. Таны тарьсан үр ирэх 9 жилийн турш ургана гэдгийг санаарай. Зоригтой бай.",
    2: "Хамтын Ажиллагаа ба Тэвчээр. Энэ жил аливаа зүйл бага зэрэг удаашралтай мэт санагдаж магадгүй ч, энэ нь үндсээ бэхжүүлэх цаг юм. Ганцаараа зүтгэх биш, бусадтай хамтрах жил. Гэрээ хэлцэл, түншлэлд анхаар. Мөнгөө эрсдэлтэй зүйлд хийхээс зайлсхийж, хадгаламжаа зузаатга. Дипломат харилцаа танд мөнгө авчирна. Тэвчээртэй байж, жижиг зүйлсэд анхаар.",
    3: "Бүтээлч Байдал ба Тэлэлт. Нийгмийн идэвх сэргэж, өөрийгөө илэрхийлэх, олон танилтай болох жил. Эрч хүчээр дүүрэн, хөгжилтэй цаг үе. Маркетинг, PR, олон нийттэй харилцах ажилд амжилт олно. Гэхдээ сэтгэл хөдлөлөөр их мөнгө үрэх, тансаг хэрэглээнд хэт автах эрсдэлтэй тул санхүүгийн сахилга баттай байгаарай.",
    4: "Хөдөлмөр ба Суурь Бэхжүүлэх. Энэ жил танаас шаргуу хөдөлмөр, эмх цэгц шаардана. 'Ажил, ажил, бас дахин ажил' гэсэн уриатай жил. Тогтвортой байдлыг бий болгоно. Үл хөдлөх хөрөнгө худалдаж авах, байшин барих, урт хугацааны төлөвлөгөө гаргахад таатай. Эрсдэлтэй алхам хийх хэрэггүй. Зөвхөн бодит, баталгаатай зүйлд мөнгөө хий.",
    5: "Өөрчлөлт ба Эрх Чөлөө. Гэнэтийн өөрчлөлтүүдээр дүүрэн жил. Аялал зугаалга, нүүдэл суудал, ажлын байраа солих зэрэг хөдөлгөөн ихтэй байна. Шинэ зах зээл, шинэ боломжуудыг эрэлхийл. Худалдаа наймаа, гадаад харилцаанд ашигтай. Гэхдээ мөнгө орж ирэхийн хэрээр хурдан гарах магадлалтай тул урсгалаа хянаарай.",
    6: "Хариуцлага ба Гэр Бүл. Гэр бүл, орон гэр, хайр дурлалын асуудал нэгдүгээрт тавигдана. Бусдын өмнө хариуцлага хүлээх, үйлчилгээ үзүүлэх жил. Гэрээ тохижуулах, гэр бүлийн гишүүдэд туслахад зарлага гарч магадгүй. Гэхдээ тогтвортой, найдвартай эх үүсвэрээс орлого орно. Урлаг, гоо сайхан, үйлчилгээний салбарт ээлтэй.",
    7: "Дотоод Шинжилгээ ба Суралцах. Өөрийгөө сонсох, ганцаараа байх, суралцах, судалгаа хийх жил. Материаллаг ертөнцөөс түр хөндийрч, оюун санаагаа цэнэглэх үе. Энэ жил бизнесээ огцом тэлэх гэж зүтгээд хэрэггүй. Түүний оронд мэргэжлийн ур чадвараа дээшлүүлж, ирээдүйн том төсөлдөө бэлд. Мөнгөний хойноос хөөцөлдөх тусам холдоно, харин мэдлэг хөөвөл мөнгө дагана.",
    8: "Эрх Мэдэл ба Ургац Хураалт. Санхүүгийн хувьд хамгийн хүчирхэг жил. Таны өнгөрсөн жилүүдэд хийсэн хөдөлмөрийн үр шим, шагнал ирнэ. Карьерын оргил үе. Том хэмжээний гэрээ хийх, бизнесээ өргөжүүлэх, хөрөнгө оруулалт хийхэд хамгийн тохиромжтой. Та удирдлагыг гартаа авч, зоригтой алхам хийх ёстой. Мөнгөний урсгал дээд цэгтээ хүрнэ.",
    9: "Төгсгөл ба Цэвэрлэгээ. 9 жилийн мөчлөгийн сүүлийн жил. Дуусгаагүй ажлаа дуусгах, хэрэггүй зүйлсээс (муу зуршил, ашиггүй харилцаа) салах жил. Шинэ мөчлөгт бэлдэх үе. Өр зээлээ барагдуулж, тооцоогоо цэгцэл. Шинэ том бизнес эхлүүлэхэд тийм ч таатай биш. Харин буяны үйл хийх, бусдад өгөх нь ирэх жилүүдийн замыг засдаг."
  },

  REMEDY_MAP: {
    "Ус": { metaphor: "Ус хэт ихдвэл Үер болж, бүхнийг сүйтгэх эрсдэлтэй. Үерийг зөвхөн Мод л шингээж, Өсөлт болгон хувиргадаг.", usage: "Ногоон болон Цайвар цэнхэр өнгөнүүд. Мөнгөө хадгалдаг саваа ногоон эсвэл гүн цэнхэр өнгөтэй болго. Ажлын ширээн дээрээ амьд ургамал тавь." },
    "Гал": { metaphor: "Гал хэт ихдвэл түймэр болно. Галыг зөвхөн Шороо буюу Газар л тогтоон барьж, Тулга болж өгдөг.", usage: "Шар, Бор, Шаргал зэрэг шороон өнгөнүүд. Түрийвчээ бор арьсан эсвэл шаргал өнгөтэй сонго. Ажлын өрөөндөө шавар эдлэл, байгалийн чулуу байршуул." },
    "Мод": { metaphor: "Мод хэт ихдвэл ширэнгэ ой болж, нарны гэрлийг хаадаг. Модны энергийг Гал буюу Илч гэрэл л идэвхжүүлдэг.", usage: "Улаан, Ягаан, Улбар шар өнгөнүүд. Мөнгөө улаан туяатай түрийвчинд хадгал. Гэрийнхээ зүүн урд зүгт тод гэрэлтүүлэг, улаан чимэглэл тавь." },
    "Төмөр": { metaphor: "Төмөр бол хатуу энерги. Ус байхгүй бол төмөр зэвэрч хугардаг. Ус л төмрийг угааж Урсгал оруулдаг.", usage: "Хар, Хөх, Гүн цэнхэр өнгөнүүд. Хар өнгийн түрийвч хамгийн тохиромжтой. Ажлын ширээн дээрээ жижиг усан оргилуур эсвэл тунгалаг шилэн эдлэл байршуул." },
    "Шороо": { metaphor: "Шороо нь дотор нь нуугдсан Төмөр буюу Алт Эрдэнэсээ ил гаргах үед л үнэ цэнэтэй болдог.", usage: "Цагаан, Саарал, Мөнгөлөг, Алтлаг өнгөнүүд. Цагаан болон мөнгөлөг өнгийн түрийвч хэрэглэ. Биедээ алт, мөнгөн гоёл чимэглэл, цаг зүү." }
  }
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
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return; 

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
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
        console.log(`Processing Astrology logic for user...`);
        
        const profile = parseAndCalculateProfile(inputData, KEYS.GEMINI);
        const reportResult = generateSequentialReport(profile, KEYS.GEMINI);
        
        // 🛡️ Safe PDF Generation (Smart Insertion)
        const pdfUrl = createPdfSafely(name, reportResult.text, KEYS.TEMPLATE, KEYS.FOLDER);
        
        sendUChatProven(contactID, pdfUrl, name, KEYS.UCHAT);

        const totalTokens = (profile.parsingUsage || 0) + reportResult.usage;
        const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");

        pdfCell.setValue(pdfUrl); 
        statusCell.setValue("АМЖИЛТТАЙ"); 
        tokenCell.setValue(totalTokens);
        debugCell.setValue(JSON.stringify(profile)); 
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
        } else if (errorMsgStr.includes("Drive")) {
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
// 2. ASTROLOGY LOGIC (VERIFIED)
// ==========================================
function parseAndCalculateProfile(rawInput, apiKey) {
  const normalized = normalizeInputWithAI(rawInput, apiKey);
  const dateStr = normalized.date; 
  const gender = normalized.gender; 
  const name = normalized.name || "Эрхэм";
  const [year, month, day] = dateStr.split(".").map(Number);
  
  const lifePath = calculateLifePath(year, month, day);
  const mongolData = getMongolianYearData(year, month, day);
  const elementRaw = mongolData.elementRaw; 
  const elementName = CONFIG.ELEMENT_MAP[elementRaw];
  const missingNumbers = calculateMissingNumbers(year, month, day);
  
  const py1 = calculatePersonalYear(month, day, 2026);
  const py2 = calculatePersonalYear(month, day, 2027);
  const py3 = calculatePersonalYear(month, day, 2028);
  
  const remedy = CONFIG.REMEDY_MAP[elementRaw];
  const lpName = CONFIG.LIFE_PATH_MAP[lifePath.number] || "Unknown";
  
  const allDigits = `${year}${month}${day}`.split("").join(", ");
  const gridDisplay = `[ ${allDigits} ]`;

  const missingText = missingNumbers.length > 0 
    ? missingNumbers.map(n => `**${n}-ийн тоо дутуу:** ${CONFIG.MISSING_NUM_MAP[n]}`).join("\n\n")
    : "Таны кодонд бүх тоо тэнцвэртэй байна.";

  return {
    dob: dateStr, year: year, gender: gender, name: name,
    parsingUsage: normalized.usage,
    lifePath: lifePath.number, lifePathCalc: lifePath.calculation, lifePathName: lpName,
    element: elementName, rawElement: elementRaw,
    missingNumbers: missingNumbers.join(", "), missingNumbersText: missingText, gridDisplay: gridDisplay,
    py1: py1, py1Text: CONFIG.PERSONAL_YEAR_MAP[py1],
    py2: py2, py2Text: CONFIG.PERSONAL_YEAR_MAP[py2],
    py3: py3, py3Text: CONFIG.PERSONAL_YEAR_MAP[py3],
    remedyMetaphor: remedy.metaphor, remedyUsage: remedy.usage
  };
}

function getMongolianYearData(year, month, day) {
  const tsDate = CONFIG.TSAGAAN_SAR[year];
  if (!tsDate) return { elementRaw: CONFIG.ELEMENTS_BY_LAST_DIGIT[year % 10] };
  const [tsMonth, tsDay] = tsDate.split("-").map(Number);
  let trueYear = year;
  if (month < tsMonth || (month === tsMonth && day < tsDay)) trueYear = year - 1;
  return { elementRaw: CONFIG.ELEMENTS_BY_LAST_DIGIT[trueYear % 10] };
}

function calculateMissingNumbers(y, m, d) {
  const present = new Set(`${y}${m}${d}`.split("").map(Number));
  return [1,2,3,4,5,6,7,8,9].filter(i => !present.has(i));
}

function calculatePersonalYear(m, d, currentYear) {
  const sumDigits = (n) => String(n).split('').reduce((a, b) => a + Number(b), 0);
  let total = sumDigits(currentYear) + sumDigits(m) + sumDigits(d);
  while (total > 9) total = sumDigits(total);
  return total;
}

function calculateLifePath(y, m, d) {
  const sumDigits = (n) => String(n).split('').reduce((a, b) => a + Number(b), 0);
  const dateString = `${y}${m < 10 ? '0'+m : m}${d < 10 ? '0'+d : d}`;
  let total = dateString.split("").reduce((a, b) => a + Number(b), 0);
  const initialSum = total;
  while (total > 9 && total !== 11 && total !== 22 && total !== 33) total = sumDigits(total);
  return { number: total, calculation: `${y}.${m}.${d} -> ${initialSum} -> ${total}` };
}

function normalizeInputWithAI(raw, key) {
  const prompt = `
  TASK: Normalize this input data for an Astrology App.
  INPUT: "${raw}"
  INSTRUCTIONS:
  1. Extract Date: Convert to YYYY.MM.DD.
  2. Extract Gender: "Эрэгтэй"/"Эмэгтэй". Default "Эмэгтэй".
  3. Extract Name: If present, format properly.
  
  RETURN JSON ONLY: {"date": "YYYY.MM.DD", "gender": "Эрэгтэй", "name": "Name"}
  `;
  try {
    const result = callGeminiAPI(prompt, key, 0.1, true);
    const data = JSON.parse(result.text.trim());
    if (!data.date) throw new Error("No date");
    
    let gender = "Эмэгтэй";
    let rawG = (data.gender || "").toLowerCase();
    if (rawG.includes("эр") || rawG === "male" || rawG === "man") gender = "Эрэгтэй";
    
    return { ...data, gender: gender, usage: result.usage };
  } catch (e) {
    const dateMatch = raw.match(/(\d{4})[\.\-\s\/](\d{1,2})[\.\-\s\/](\d{1,2})/);
    let fallbackGender = "Эмэгтэй";
    if (raw.toLowerCase().includes("эр")) fallbackGender = "Эрэгтэй";
    return {
      date: dateMatch ? `${dateMatch[1]}.${dateMatch[2]}.${dateMatch[3]}` : "2000.01.01",
      gender: fallbackGender, name: "Эрхэм", usage: 0
    };
  }
}

// ==========================================
// 3. AI GENERATION (TOKEN-SAFE LENGTH + STRICT FORMAT)
// ==========================================
function generateSequentialReport(data, apiKey) {
  const SYSTEM_PROMPT = `
  ROLE: Astrologer & Wise Mentor (Зурхайч).
  TONE: Calm, Respectful, Traditional, Spiritual, Deep.
  LANGUAGE: Proper Mongolian Cyrillic ONLY.
  
  >>> ХАТУУ ДҮРМҮҮД (ШУУД МӨРДӨХ): <<<
  1. УРТ БОЛОН ДЭЛГЭРЭНГҮЙ БАЙДАЛ: Тайлан нь дэлгэрэнгүй, гүнзгий, шинжлэх ухаанч байх ёстой. Энэ хэсэг нь НИЙТДЭЭ 800 орчим үгтэй байх шаардлагатай. Хэт нуршихгүйгээр, гол санааг утга төгөлдөр тайлбарлаж өгүүлбэрээ бүрэн гүйцээж дуусга.
  2. ЭМОЖИНЫ ДҮРЭМ: Догол мөр бүрийн зөвхөн хамгийн эхэнд 1 ширхэг эможи (жишээ нь: 🌓, 💡, 🌑, ⚖️, 🔍, ⚡️, ⏳) тавина. Өгүүлбэр дунд эсвэл төгсгөлд эможи огт ашиглахыг ХОРИГЛОНО!
  3. БҮТЭЦ: Од (*) эсвэл зураас (-) ашиглан хэсгүүдийг салгахыг хориглоно. Зөвхөн догол мөрөөр тусгаарла.
  `;

  const promptBatch1 = `
  ${SYSTEM_PROMPT}
  TASK: Write PART 1 and PART 2 together.
  
  DATA FOR PART 1:
  - DOB: ${data.dob}
  - Life Path: ${data.lifePath} (${data.lifePathName}) [Calc: ${data.lifePathCalc}]
  - Element: ${data.element}
  START PART 1 WITH EXACTLY: "1-р хэсэг: ТАНЫ ЗАЯА ТӨӨРГӨӨС ӨГӨГДСӨН ХИШИГ"

  DATA FOR PART 2:
  - Light/Shadow of ${data.lifePathName} and ${data.element}.
  START PART 2 WITH EXACTLY: "2-р хэсэг: ТАНЫ ЗАН ЧАНАРЫН ГЭГЭЭ БА СҮҮДЭР"
  `;
  const r1 = callGeminiAPI(promptBatch1, apiKey, CONFIG.TEMPERATURE);

  const promptBatch2 = `
  ${SYSTEM_PROMPT}
  TASK: Write PART 3 and PART 4 together.
  
  DATA FOR PART 3:
  - Visual Grid: ${data.gridDisplay}
  - Missing Numbers Text: ${data.missingNumbersText}
  START PART 3 WITH EXACTLY: "3-р хэсэг: ЭД БАЯЛГИЙГ ТОГТООХ САВНЫ ЦООРХОЙ"
  
  DATA FOR PART 4:
  - 2026: Personal Year ${data.py1} -> "${data.py1Text}"
  - 2027: Personal Year ${data.py2} -> "${data.py2Text}"
  - 2028: Personal Year ${data.py3} -> "${data.py3Text}"
  START PART 4 WITH EXACTLY: "4-р хэсэг: ИРЭЭДҮЙ ЦАГИЙН ОДОД (2026-2028)"
  `;
  const r2 = callGeminiAPI(promptBatch2, apiKey, CONFIG.TEMPERATURE);

  const promptBatch3 = `
  ${SYSTEM_PROMPT}
  TASK: Write PART 5.
  
  DATA:
  - Element: ${data.rawElement}
  - Metaphor: ${data.remedyMetaphor}
  - Usage: ${data.remedyUsage}
  
  START PART 5 WITH EXACTLY: "5-р хэсэг: ЗАЯАГАА ТЭГШЛЭХ ЗАСАЛ БА ЕРӨӨЛ"
  `;
  const r3 = callGeminiAPI(promptBatch3, apiKey, CONFIG.TEMPERATURE);

  return {
    text: r1.text.trim() + "\n\n" + r2.text.trim() + "\n\n" + r3.text.trim(),
    usage: (r1.usage||0) + (r2.usage||0) + (r3.usage||0)
  };
}

// ==========================================
// 4. API & SAFE PDF DELIVERY ENGINE
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

  const res = UrlFetchApp.fetch(url, { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true });
  const json = JSON.parse(res.getContentText()); 
  
  if (json.candidates && json.candidates[0].content) {
      return { 
          text: json.candidates[0].content.parts[0].text, 
          usage: json.usageMetadata ? json.usageMetadata.totalTokenCount : 0 
      };
  }
  throw new Error(`Gemini API Error: ${res.getContentText()}`);
}

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
  
  body.replaceText("(?i){{report}}", "INSERT_REPORT_HERE_999");
  
  let insertionIndex = -1;
  const searchResult = body.findText("INSERT_REPORT_HERE_999"); 
  
  if (searchResult) {
    const element = searchResult.getElement();
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
        
        if (pText.length > 50) p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY); 
        
        if (/^\d-р хэсэг:/.test(pText)) {
            p.setHeading(DocumentApp.ParagraphHeading.HEADING2);
            p.setBold(true);
            p.setSpacingBefore(15); 
        } else {
            p.setLineSpacing(1.5); 
        }
        p.setSpacingAfter(10); 
    }
  }

  doc.saveAndClose(); 

  const pdfBlob = copyFile.getAs('application/pdf');
  const pdfFile = targetFolder.createFile(pdfBlob);
  pdfFile.setName(`${name} - Зурхай.pdf`); 
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); 
  
  copyFile.setTrashed(true);
  return pdfFile.getUrl(); 
}

function sendUChatProven(userNs, pdfUrl, name, token) {
  if (!token) throw new Error("DELIVERY: UChat token байхгүй.");
  if (!userNs) throw new Error("DELIVERY: user_ns хоосон.");
  
  const msg = CONFIG.UCHAT.DELIVERY_MESSAGE.replace("{{NAME}}", name);
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
