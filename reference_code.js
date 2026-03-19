// --- КОДЫН ЭХЛЭЛ ---

/****************************************************************************************
 * PRODUCT: MONGOLIAN COUPLE ASTROLOGY (HARMONY V18.9)
 * ENGINE: ROBUST V26.1 (Smart Doc Insertion, Token Limit Safe, UChat, Regex Cleaner)
 ****************************************************************************************/

const CONFIG = {
  VERSION: "v26.1-Couple-Harmony",
  PRODUCT_NAME: "Хосын Нийцлийн Алтан Тайлан",
  SHEET_NAME: "Sheet1",
  BATCH_SIZE: 3, 
  GEMINI_MODEL: "gemini-2.0-flash",
  TEMPERATURE: 0.3, 

  COLUMNS: {
    NAME: 0, ID: 1, INPUT: 2, PDF: 3, STATUS: 4, 
    TOKEN: 5, DEBUG: 6, DATE: 7, VER: 8, ERROR: 9
  },

  UCHAT: {
    ENDPOINT: "https://www.uchat.com.au/api/subscriber/send-content",
    DELIVERY_MESSAGE: `Сайн байна уу, {{NAME}}? 🔮\n\nТа хоёрын "Хосын Нийцлийн Алтан Тайлан" бэлэн боллоо.\n\nДоорх товч дээр дарж татаж авна уу. 👇`,
    DELIVERY_BTN_TEXT: "📥 Тайлан татах"
  },

  // ==================================================================================
  // 📜 FULL "GOLD STANDARD" REFERENCE TEXTS (STYLE GUIDE ONLY)
  // ==================================================================================
  REFERENCES: {
    PART_1: `1-р хэсэг. ЯЗГУУР МӨН ЧАНАР: Жилүүдийн Ивээл ба Харш
Эхлээд та хоёрын төрсөн жилийн мөн чанарыг тодорхойлбол... Монгол зурхайн гүн ухаанд хоёр амьтан нь хоорондоо ямар барилдлагатай байдаг тул суурь харилцаа нь ямар байхыг илтгэнэ. Тэдний харилцаанд сүүдэр тусах уу, амар амгалан зонхилох уу.
🐉 Эрэгтэй нь [Амьтан] жилтэй тул байгалийн өгөгдөл нь... (амьтны онцлог, манлайлагч эсвэл нямбай чанар). Түүний [Махбод] нь түүнийг илүү шийдэмгий, хатуужилтай эсвэл уян хатан гэдгийг харуулна.
🐔 Эмэгтэй нь [Амьтан] жилтэй тул төрөлхийн... (амьтны онцлог). Түүний [Махбод] нь түүнийг ямар болгож, хурц занг нь хэрхэн зөөлрүүлдэг эсвэл хурцалдаг.
💖 Та хоёрын жилийн махбодын харилцааг шинжвэл маш онцгой зохицол эсвэл харш харагдаж байна. Учир нь эрэгтэйн махбод нь эмэгтэйн махбодыг төрүүлдэг үү, эсвэл сүйтгэдэг үү. Үүнийг зурхайн ухаанд юу гэж үздэг бөгөөд энэ нь бодит амьдрал дээр юу гэсэн үг юм.
🔑 Дүгнэж хэлэхэд [Амьтан], [Амьтан] хоёр нь жилийн хувьд ямар барилдлагатай бөгөөд махбодын хувьд ямар тул та хоёрын хоорондын язгуур мөн чанарын зохицол ямар байна. Гол анхаарах зүйл нь юу вэ.`,

    PART_2: `2-р хэсэг. ЭРЧИМ ХҮЧНИЙ НИЙЛЭМЖ: Махбодын Зохицол
Энэ хэсэгт бид та хоёрын зөвхөн төрсөн жил бус, тухайн жилийг удирдаж буй байгалийн суурь энерги буюу "Махбод"-ын харилцааг гүнзгийрүүлэн авч үзнэ.
💧 Эрэгтэй нь [Махбод], эмэгтэй нь [Махбод] байгаа нь энергийн солилцооны хувьд "Төрүүлэгч ба Төрөгдөгч" эсвэл "Харш" гэсэн урсгалыг үүсгэдэг. Яагаад гэвэл байгаль дээр энэ хоёр махбод хэрхэн харилцдаг тул амьдрал дээр бие биенээ хэрхэн тэтгэдэг эсвэл дардаг.
🌊 Сэтгэл хөдлөлийн хэмнэлийн хувьд та хоёр бие биеэ хэрхэн нөхөх вэ. Эрэгтэй нь ямар хариу үйлдэл үзүүлж, эмэгтэй нь яаж зөөлрүүлэх эсвэл хурцатгах вэ.
🩺 Эрүүл мэндийн болон бие махбодын нийцлийн тухайд энэ хоёр махбод нь бие биедээ "эм" эсвэл "хор" болдог эсэх. Аль эрхтэн тогтолцоог тэтгэх эсвэл ядраах.
🌱 Эцэст нь дүгнэхэд, энэхүү махбодын зохицол нь та хоёрын харилцааг ямар болгож байна. Ганцхан анхаарах зүйл нь юу вэ.`,

    PART_3: `3-р хэсэг. ОЮУН САНААНЫ ХОЛБООС: Мэнгэний Таарамж
🧠 Энэ хэсэгт та хоёрын сэтгэлгээний онцлог, амьдралыг үзэх үзэл болон оюун санааны нийцлийг "Мэнгэ"-ний зурхайгаар тайлбарлая. Тооцоолж үзвэл эрэгтэй нь [Мэнгэ], эмэгтэй нь [Мэнгэ] байна. Энэ нийлэмж нь Монгол зурхайд ямар барилдлага гэж буудаг.
🔥 Эрэгтэйн мэнгэ нь байгалийн мөн чанараараа ямар махбодыг илтгэж, оюун санааны хувьд ямар онцлогтой байх вэ.
⚔️ Эмэгтэйн мэнгэ нь ямар махбодыг илтгэж, оюун санааны хувьд ямар зарчимтай байх вэ.
💥 Та хоёрын харилцааг ингэж нэрлэсний учир нь юу вэ? Оюун санааны хувьд бие биетэйгээ хэрхэн мөргөлдөх эсвэл нийлж нэгдэх вэ?
🤝 Энэ харилцаанд гүн утга учир бий. Та хоёр бие биенийхээ сул талыг хэрхэн зөвөөр хандаж, бие биеэ хэрхэн хөгжүүлэх вэ.`,

    PART_4: `4-р хэсэг. АМЬДРАЛЫН ХЭМНЭЛ: Суудлын Өнцөг & Гал Голомт
🏠 Энэ хэсэг нь та хоёрын өдөр тутмын ахуй амьдрал, гэр орны уур амьсгал, санхүүгийн удирдлага болон гал тогоонд хэрхэн зохицохыг харуулна.
🔥 Эрэгтэй нь [Суудал] суудалтай. Энэ нь ямар зүг, улирал, бэлгэдэлтэй вэ. Ийм суудалтай хүн гэр бүлдээ ямар хэв маягийг авчирдаг.
⛰️ Эмэгтэй нь [Суудал] суудалтай. Энэ нь гэр бүлдээ ямар тогтвортой байдал эсвэл хөдөлгөөнийг авчирдаг вэ.
⚠️ Суудлын нийцлийг шинжвэл ямар харилцаа үүсэж байна. Энэ нь амьдрал дээр яг юу гэсэн үг вэ? Санхүүгийн тал дээр хэн нь үрж, хэн нь хадгалах гэх мэт.
☯️ Гэхдээ бүү санаа зов, Зурхай бол онош, Дом бол эмчилгээ. Та хоёрын энэ зөрчлийг арилгахын тулд гэрийнхээ уур амьсгалд ямар өнгө, бэлгэдлийг тэнцвэртэй оруулах хэрэгтэй вэ (Засал).`,

    PART_5: `5-р хэсэг. ҮЙЛИЙН ҮР БА ИРЭЭДҮЙН ЗУРАГЛАЛ Нэгдсэн Дүгнэлт
🔮 Та хоёрын харилцааг нэг үгээр тодорхойлбол... (үйлийн үрээр учирсан, эсвэл бие биеэ нөхөх холбоо). Яагаад уулзсан бэ?
🔗 Ирээдүйд юуг анхаарах вэ? Ямар сорилтууд тулгарах вэ? Тэдгээр сорилтууд нь та хоёрыг салгах биш, улам бэхжүүлэх боломж гэж харах хэрэгтэй.
️ Та хоёрын аз жаргалын түлхүүр буюу дом нь юу вэ? Таван махбодыг харилцаандаа хэрхэн ухамсартайгаар тэнцвэржүүлэх вэ.
💎 Эцсийн үнэлгээгээр та хоёрын харилцаа хэдэн хувьтай байна. Хамтдаа уулсыг ч нүүлгэж чадах уу? Сайн сайхныг ерөөе!`
  },

  TSAGAAN_SAR_MAP: {
    1940: "02-08", 1941: "01-27", 1942: "02-15", 1943: "02-05", 1944: "02-24",
    1945: "02-13", 1946: "02-02", 1947: "01-22", 1948: "02-10", 1949: "01-29",
    1950: "02-17", 1951: "02-06", 1952: "01-27", 1953: "02-14", 1954: "02-03",
    1955: "02-24", 1956: "02-12", 1957: "01-31", 1958: "02-18", 1959: "02-08",
    1960: "02-27", 1961: "02-15", 1962: "02-05", 1963: "02-25", 1964: "02-13",
    1965: "02-02", 1966: "02-21", 1967: "02-09", 1968: "01-30", 1969: "02-17",
    1970: "02-06", 1971: "02-27", 1972: "02-15", 1973: "02-06", 1974: "02-23",
    1975: "02-11", 1976: "01-31", 1977: "02-18", 1978: "02-07", 1979: "02-28",
    1980: "02-16", 1981: "02-05", 1982: "02-24", 1983: "02-13", 1984: "02-02",
    1985: "02-20", 1986: "02-09", 1987: "01-29", 1988: "02-17", 1989: "02-06",
    1990: "02-27", 1991: "02-15", 1992: "02-04", 1993: "02-23", 1994: "02-10",
    1995: "01-31", 1996: "02-19", 1997: "02-07", 1998: "02-28", 1999: "02-16",
    2000: "02-05", 2001: "02-24", 2002: "02-12", 2003: "02-01", 2004: "02-22",
    2005: "02-09", 2006: "01-29", 2007: "02-18", 2008: "02-07", 2009: "02-25",
    2010: "02-14", 2011: "02-03", 2012: "02-22", 2013: "02-11", 2014: "01-31",
    2015: "02-19", 2016: "02-09", 2017: "02-27", 2018: "02-16", 2019: "02-05",
    2020: "02-24", 2021: "02-12", 2022: "02-02", 2023: "02-21", 2024: "02-10",
    2025: "02-28" 
  },

  // 1. ANIMAL & ELEMENT LOGIC
  ANIMALS: ["Бич", "Тахиа", "Нохой", "Гахай", "Хулгана", "Үхэр", "Бар", "Туулай", "Луу", "Могой", "Морь", "Хонь"], 
  
  ANIMAL_MAP: [
    "Хулгана", "Үхэр", "Бар", "Туулай", "Луу", "Могой", 
    "Морь", "Хонь", "Бич", "Тахиа", "Нохой", "Гахай"
  ],

  ELEMENT_MAP: {
    0: "Төмөр", 1: "Төмөр", 
    2: "Ус", 3: "Ус", 
    4: "Мод", 5: "Мод", 
    6: "Гал", 7: "Гал", 
    8: "Шороо", 9: "Шороо"
  },

  // 2. MENGE MAPPING (1-9)
  MENGE_NAMES: {
    1: "1 цагаан", 2: "2 хар", 3: "3 хөх", 4: "4 ногоон", 
    5: "5 шар", 6: "6 цагаан", 7: "7 улаан", 8: "8 цагаан", 9: "9 улаан"
  },
  
  MENGE_ELEMENTS: {
    1: "Ус", 2: "Шороо", 3: "Мод", 4: "Мод", 
    5: "Шороо", 6: "Төмөр", 7: "Төмөр", 8: "Шороо", 9: "Гал"
  },

  // 3. SUUDAL MAPPING (Trigrams)
  SUUDAL_NAMES: {
    1: "Хам", 2: "Хон", 3: "Зэн", 4: "Цон", 
    6: "Төмөр", 7: "Дүй", 8: "Гэн", 9: "Ли"
  }, 

  SUUDAL_ELEMENTS: {
    "Хам": "Ус", "Хон": "Шороо", "Зэн": "Мод", "Цон": "Мод", 
    "Төмөр": "Төмөр", "Дүй": "Төмөр", "Гэн": "Шороо", "Ли": "Гал"
  },

  // 4. MATRIX LOGIC (The "Brain")
  ELEMENT_RELATIONS: {
    "Гал-Шороо": "Эх хүү (Төрүүлэх)", "Шороо-Гал": "Эх хүү (Төрүүлэх)",
    "Шороо-Төмөр": "Эх хүү (Төрүүлэх)", "Төмөр-Шороо": "Эх хүү (Төрүүлэх)",
    "Төмөр-Ус": "Эх хүү (Төрүүлэх)", "Ус-Төмөр": "Эх хүү (Төрүүлэх)",
    "Ус-Мод": "Эх хүү (Төрүүлэх)", "Мод-Ус": "Эх хүү (Төрүүлэх)",
    "Мод-Гал": "Эх хүү (Төрүүлэх)", "Гал-Мод": "Эх хүү (Төрүүлэх)",
    
    "Гал-Гал": "Ижил нөхөр", "Шороо-Шороо": "Ижил нөхөр",
    "Төмөр-Төмөр": "Ижил нөхөр", "Ус-Ус": "Ижил нөхөр", "Мод-Мод": "Ижил нөхөр",
    
    "Ус-Гал": "Өшөөт дайсан (Харш)", "Гал-Ус": "Өшөөт дайсан (Харш)",
    "Гал-Төмөр": "Өшөөт дайсан (Харш)", "Төмөр-Гал": "Өшөөт дайсан (Харш)",
    "Төмөр-Мод": "Өшөөт дайсан (Харш)", "Мод-Төмөр": "Өшөөт дайсан (Харш)",
    "Мод-Шороо": "Өшөөт дайсан (Харш)", "Шороо-Мод": "Өшөөт дайсан (Харш)",
    "Шороо-Ус": "Өшөөт дайсан (Харш)", "Ус-Шороо": "Өшөөт дайсан (Харш)"
  },

  MENGE_OUTCOMES: {
    "Healer": "Эмч домч",
    "Best": "Их Ивээл",
    "Good": "Бага Ивээл",
    "Bad": "Харш"
  },

  SUUDAL_GOOD_PAIRS: {
    "Хам-Хам": "Сан", "Хон-Хон": "Сан", "Зэн-Зэн": "Сан", "Цон-Цон": "Сан",
    "Төмөр-Төмөр": "Сан", "Дүй-Дүй": "Сан", "Гэн-Гэн": "Сан", "Ли-Ли": "Сан",
    
    "Хам-Ли": "Билэгүүн", "Ли-Хам": "Билэгүүн",
    "Зэн-Цон": "Билэгүүн", "Цон-Зэн": "Билэгүүн",
    "Хон-Төмөр": "Билэгүүн", "Төмөр-Хон": "Билэгүүн",
    "Гэн-Дүй": "Билэгүүн", "Дүй-Гэн": "Билэгүүн",
    
    "Хам-Зэн": "Эмч", "Зэн-Хам": "Эмч",
    "Цон-Ли": "Эмч", "Ли-Цон": "Эмч",
    "Хон-Дүй": "Эмч", "Дүй-Хон": "Эмч",
    "Гэн-Төмөр": "Эмч", "Төмөр-Гэн": "Эмч",
    
    "Хам-Цон": "Огторгуй", "Цон-Хам": "Огторгуй",
    "Зэн-Ли": "Огторгуй", "Ли-Зэн": "Огторгуй",
    "Хон-Гэн": "Огторгуй", "Гэн-Хон": "Огторгуй",
    "Дүй-Төмөр": "Огторгуй", "Төмөр-Дүй": "Огторгуй"
  },

  SUUDAL_BAD_PAIRS: {
    "Хам-Хон": "Харш", "Хон-Хам": "Харш",
    "Зэн-Дүй": "Харш", "Дүй-Зэн": "Харш",
    "Цон-Гэн": "Харш", "Гэн-Цон": "Харш",
    "Ли-Төмөр": "Харш", "Төмөр-Ли": "Харш",
    
    "Хам-Гэн": "Адид", "Гэн-Хам": "Адид",
    "Зэн-Төмөр": "Адид", "Төмөр-Зэн": "Адид",
    "Цон-Хон": "Адид", "Хон-Цон": "Адид",
    "Ли-Дүй": "Адид", "Дүй-Ли": "Адид",
    
    "Хам-Дүй": "Хорт", "Дүй-Хам": "Хорт",
    "Зэн-Хон": "Хорт", "Хон-Зэн": "Хорт",
    "Цон-Төмөр": "Хорт", "Төмөр-Цон": "Хорт",
    "Ли-Гэн": "Хорт", "Гэн-Ли": "Хорт",
    
    "Хам-Төмөр": "Гай", "Төмөр-Хам": "Гай",
    "Зэн-Гэн": "Гай", "Гэн-Зэн": "Гай",
    "Цон-Дүй": "Гай", "Дүй-Цон": "Гай",
    "Ли-Хон": "Гай", "Хон-Ли": "Гай"
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
  const START_TIME = new Date().getTime();
  const TIME_LIMIT_MS = 5 * 60 * 1000; // 5.0 minutes

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
        console.log(`Processing Couple Harmony logic for user...`);
        
        // 1. DATA PREP: AI PARSING
        const normalized = normalizeInputWithAI(inputData, KEYS.GEMINI);
        
        // 2. PARSE & CALCULATE (The Brain)
        const coupleData = calculateCoupleProfile(normalized);
        
        // 3. GENERATE REPORT (The Voice) - Split Calls
        const reportResult = generateSequentialReport(coupleData, KEYS.GEMINI);
        
        // 4. SAFE PDF DELIVERY ENGINE
        const pdfUrl = createPdfSafely(name, reportResult.text, KEYS.TEMPLATE, KEYS.FOLDER);
        
        sendUChatProven(contactID, pdfUrl, name, KEYS.UCHAT);

        const totalTokens = (normalized.usage || 0) + reportResult.usage;
        const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");

        const outputRange = sheet.getRange(i + 1, COLS.PDF + 1, 1, 7);
        outputRange.setValues([[pdfUrl, "АМЖИЛТТАЙ", totalTokens, JSON.stringify(coupleData), now, CONFIG.VERSION, ""]]);
        
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
// 2. CALCULATIONS ENGINE (Tsagaan Sar + Year Menge + Life Gua)
// ==========================================
function sumDigits(num) {
  return String(num).split('').reduce((a, b) => a + Number(b), 0);
}

function calculateAstroProfile(dob, gender, name) {
  const parts = dob.split(/[.\-\/]/).map(Number);
  let y = parts[0];
  const m = parts[1];
  const d = parts[2];
  
  // --- LUNAR NEW YEAR CHECK (Tsagaan Sar Logic) ---
  if (CONFIG.TSAGAAN_SAR_MAP[y]) {
    const tsagaanSar = CONFIG.TSAGAAN_SAR_MAP[y].split("-").map(Number);
    if (m < tsagaanSar[0] || (m === tsagaanSar[0] && d < tsagaanSar[1])) {
      y = y - 1; 
    }
  }

  // Animal & Element (Based on adjusted lunar year)
  // FIXED FORMULA: 1900 was Rat. 1900 % 12 = 4.  (y - 1900) % 12 maps to ANIMAL_MAP.
  // 1900 % 12 is 4.  (1900 - 1900) % 12 = 0 -> Rat.
  const animalIdx = (y - 1900) % 12;
  const animal = CONFIG.ANIMAL_MAP[animalIdx < 0 ? animalIdx + 12 : animalIdx];
  const element = CONFIG.ELEMENT_MAP[y % 10];
  
  // --- PART 3: YEAR MENGE (Standard Formula: 11 - Sum) ---
  let n = sumDigits(y);
  while (n > 9) n = sumDigits(n);
  
  let mengeNum = 11 - n;
  if (mengeNum <= 0) mengeNum += 9;
  if (mengeNum > 9) mengeNum -= 9;
  
  const mengeName = CONFIG.MENGE_NAMES[mengeNum];
  const mengeElement = CONFIG.MENGE_ELEMENTS[mengeNum];
  
  // --- PART 4: SUUDAL / LIFE GUA (Life Gua Logic) ---
  // FIXED FORMULA FOR SUUDAL:
  let suudal = "";
  if (gender === "Эрэгтэй") {
     let sumYear = sumDigits(y);
     while(sumYear > 9) sumYear = sumDigits(sumYear);
     let rem = 11 - sumYear;
     if (rem === 5) rem = 2; // Хон (Shooro)
     suudal = CONFIG.SUUDAL_NAMES[rem];
  } else {
     let sumYear = sumDigits(y);
     while(sumYear > 9) sumYear = sumDigits(sumYear);
     let rem = sumYear + 4;
     while(rem > 9) rem = sumDigits(rem);
     if (rem === 5) rem = 8; // Гэн (Shooro)
     suudal = CONFIG.SUUDAL_NAMES[rem];
  }
  const suudalElement = CONFIG.SUUDAL_ELEMENTS[suudal];
  
  return { dob, year: y, gender, name, animal, element, menge: mengeNum, mengeName, mengeElement, suudal, suudalElement };
}

function getMengeRelation(p1, p2) {
  const e1 = p1.mengeElement;
  const e2 = p2.mengeElement;
  if ((e1 === "Мод" && e2 === "Шороо") || (e1 === "Шороо" && e2 === "Мод")) return CONFIG.MENGE_OUTCOMES.Healer;
  const feed = CONFIG.ELEMENT_RELATIONS[`${e1}-${e2}`] || "";
  if (feed.includes("Төрүүлэх")) return CONFIG.MENGE_OUTCOMES.Best;
  if (feed.includes("Ижил")) return CONFIG.MENGE_OUTCOMES.Good;
  return CONFIG.MENGE_OUTCOMES.Bad;
}

function calculateRemedy(e1, e2, relationName) {
  if (["Сан", "Билэгүүн", "Эмч", "Огторгуй"].includes(relationName)) {
    return { element: "Байхгүй", advice: "Тусгай засал шаардлагагүй, маш сайн." };
  }
  const pair = [e1, e2].sort().join("-");
  if (pair === "Мод-Төмөр") return { element: "Ус", advice: "Хар хөх өнгө, Ус тахих, Усан оргилуур" };
  if (pair === "Мод-Шороо") return { element: "Гал", advice: "Улаан өнгө, Гал тахих, Лаа асаах" };
  if (pair === "Төмөр-Шороо") return { element: "Төмөр", advice: "Цагаан өнгө, Металл эдлэл" }; 
  if (pair === "Ус-Шороо") return { element: "Төмөр", advice: "Цагаан/Мөнгөлөг өнгө, Металл эдлэл" };
  if (pair === "Гал-Төмөр") return { element: "Шороо", advice: "Шар/Бор өнгө, Керамик, Чулуу" };
  if (pair === "Гал-Ус") return { element: "Мод", advice: "Ногоон өнгө, Мод тарих" };
  return { element: "Бүх махбод", advice: "Засал хийлгэх" };
}

function calculateCoupleProfile(normalizedData) {
  const userGender = (String(normalizedData.userGender).toLowerCase() === "эрэгтэй") ? "Эрэгтэй" : "Эмэгтэй";
  const partnerGender = userGender === "Эрэгтэй" ? "Эмэгтэй" : "Эрэгтэй"; 
  
  const user = calculateAstroProfile(normalizedData.dob1, userGender, normalizedData.name);
  const partner = calculateAstroProfile(normalizedData.dob2, partnerGender, "Түнш");
  
  // --- COMPATIBILITY LOGIC ---
  const elementKey = `${user.element}-${partner.element}`;
  const elementRelation = CONFIG.ELEMENT_RELATIONS[elementKey] || "Тодорхойгүй";
  
  const mengeRelation = getMengeRelation(user, partner);
  
  const suudalKey = `${user.suudal}-${partner.suudal}`;
  let suudalRelation = CONFIG.SUUDAL_GOOD_PAIRS[suudalKey];
  if (!suudalRelation) suudalRelation = CONFIG.SUUDAL_BAD_PAIRS[suudalKey] || "Тодорхойгүй";
  
  const remedy = calculateRemedy(user.suudalElement, partner.suudalElement, suudalRelation);

  let score = 50;
  if (elementRelation.includes("Эх хүү")) score += 15;
  if (elementRelation.includes("Ижил")) score += 10;
  if (mengeRelation.includes("Их Ивээл")) score += 15;
  if (mengeRelation.includes("Эмч")) score += 20; 
  if (["Сан", "Билэгүүн", "Эмч", "Огторгуй"].includes(suudalRelation)) score += 15;
  if (score > 100) score = 100;

  return { user, partner, relations: { element: elementRelation, menge: mengeRelation, suudal: suudalRelation, remedy, score } };
}

function normalizeInputWithAI(raw, key) {
  const prompt = `
  TASK: Normalize this couple's data.
  INPUT: "${raw}"
  
  INSTRUCTIONS:
  1. Extract User's Gender (M/F). Return "Эрэгтэй" or "Эмэгтэй". Default "Эмэгтэй" if unknown.
  2. Extract User's DOB (YYYY.MM.DD).
  3. Extract Partner's DOB (YYYY.MM.DD).
  4. Use "Name" if present, else "Эрхэм".
  
  RETURN JSON ONLY: {"userGender": "Эрэгтэй", "dob1": "YYYY.MM.DD", "dob2": "YYYY.MM.DD", "name": "Name"}
  `;
  try {
    const result = callGeminiAPI(prompt, key, 0.1, true);
    const data = JSON.parse(result.text.trim());
    if (!data.dob1 || !data.dob2) throw new Error("No date");
    
    let gender = "Эмэгтэй";
    let rawG = (data.userGender || "").toLowerCase();
    if (rawG.includes("эр") || rawG === "male" || rawG === "man") gender = "Эрэгтэй";
    
    return { ...data, userGender: gender, usage: result.usage };
  } catch (e) {
    const dates = raw.match(/\d{4}[\.\-\s\/]\d{1,2}[\.\-\s\/]\d{1,2}/g) || ["1990.01.01", "1990.01.01"];
    let fallbackGender = "Эмэгтэй";
    if (raw.toLowerCase().includes("эр")) fallbackGender = "Эрэгтэй";
    
    return {
      dob1: dates[0].replace(/[\s\-\/]/g, ".") || "1990.01.01",
      dob2: (dates[1] || "1990.01.01").replace(/[\s\-\/]/g, "."),
      userGender: fallbackGender, 
      name: "Эрхэм", 
      usage: 0
    };
  }
}

// ==========================================
// 3. AI GENERATION (TOKEN-SAFE LENGTH + STRICT FORMAT)
// ==========================================
function generateSequentialReport(data, apiKey) {
  const { user, partner, relations } = data;
  let male = (user.gender === "Эрэгтэй") ? user : partner;
  let female = (user.gender === "Эмэгтэй") ? user : partner;
  
  const header = `
👨 ЭРЭГТЭЙН ӨГӨГДӨЛ: ${male.dob} | ${male.element} ${male.animal} | ${male.mengeName} | ${male.suudal}
👩 ЭМЭГТЭЙН ӨГӨГДӨЛ: ${female.dob} | ${female.element} ${female.animal} | ${female.mengeName} | ${female.suudal}
`;

  const SYSTEM_PROMPT = `
  ROLE: You are an elite, highly intuitive Mongolian Astrologer & Psychologist. Tone: Direct, Grounded, Psychological, and Wise. Avoid excessive analogies and fairy-tale similes. Write as a mature, experienced counselor speaking directly to the couple about real life.
  LANGUAGE: Proper Mongolian Cyrillic ONLY.
  
  >>> MASTER RULES (STRICTLY ENFORCED): <<<
  1. ZERO META-TALK: NEVER use phrases like "Энэ хэсэгт бид...", "Дүгнэж хэлэхэд...", "Энэ хэсэг нь...". Start your analysis IMMEDIATELY in the very first sentence.
  2. NO POETRY / FAIRY TALES: Do NOT write like a poem or a fairy tale. Be realistic, deeply psychological, and analytical. STRICTLY FORBIDDEN to use words like "мэт", "шиг" to make poetic analogies. 
  3. MONGOLIAN GRAMMAR FOR ELEMENTS: NEVER say "Ус махбодтой Тахиа". You MUST seamlessly combine them using correct grammar: "Усан тахиа", "Модон луу", "Шороон үхэр", "Төмөр бар", "Гал морь". ALWAYS use "жилтэй" instead of "амьтантай".
  4. STRICT EMOJI RULE: EVERY SINGLE PARAGRAPH (except the section title) MUST start with EXACTLY ONE emoji (e.g., 🐉, 🌊, 🔥). ZERO exceptions. Do NOT use any emojis in the middle or end of sentences.
  5. STRICT FORMATTING: 
     - NO Markdown headers like (#, ##).
     - NO Markdown bold formatting like (**text**).
     - Return ONLY PLAIN TEXT separated by double line breaks.
  6. COMPLETENESS: NEVER cut off mid-sentence. Always finish your thoughts and provide a complete, well-rounded conclusion. Aim for around 400 words, but prioritize COMPLETENESS.
  `;

  // 1st CALL: Foundation (Animals & General Elements)
  const prompt1 = `
  ${SYSTEM_PROMPT}
  TASK: Write PART 1 ONLY (The Root/Foundation).
  
  DATA: 
  - Эрэгтэй: ${male.year} оны ${male.element === 'Ус' ? 'Усан' : male.element === 'Мод' ? 'Модон' : male.element === 'Шороо' ? 'Шороон' : male.element} ${male.animal.toLowerCase()} жилтэй.
  - Эмэгтэй: ${female.year} оны ${female.element === 'Ус' ? 'Усан' : female.element === 'Мод' ? 'Модон' : female.element === 'Шороо' ? 'Шороон' : female.element} ${female.animal.toLowerCase()} жилтэй.
  
  INSTRUCTIONS: Begin exactly with "1-р хэсэг. ЯЗГУУР МӨН ЧАНАР: Жилүүдийн Ивээл ба Харш" on its own line, then double line break, then start your analysis.
  Write about their foundational connection. First paragraph: Introduce their animal years and general personality compatibility. Second paragraph: Dive deep into the Husband's psychological traits based on his animal. Third paragraph: Dive deep into the Wife's psychological traits based on her animal. Final paragraph: A practical psychological conclusion of how their personalities blend.
  CRITICAL: DO NOT talk about the "Element" interactions (like Эх хүү, Харш, who feeds whom) in this part. That belongs in Part 2. Focus 100% on the Animal/Year psychology and personalities here.
  
  STYLE GUIDE REFERENCE (Model your structure, depth, and tone exactly after this):
  ${CONFIG.REFERENCES.PART_1}
  `;
  const r1 = callGeminiAPI(prompt1, apiKey, CONFIG.TEMPERATURE);

  // 2nd CALL: Energy Flow (Chemical Interaction)
  const prompt2 = `
  ${SYSTEM_PROMPT}
  TASK: Write PART 2 ONLY (Energy & Elemental Dynamics).

  DATA:
  - Эрэгтэйн махбод: ${male.element}
  - Эмэгтэйн махбод: ${female.element}
  - Махбодын эрчим хүчний урсгал: ${male.element} ба ${female.element} хоорондын харилцаа (${relations.element})
  
  INSTRUCTIONS: Begin exactly with "2-р хэсэг. ЭРЧИМ ХҮЧНИЙ НИЙЛЭМЖ: Махбодын Зохицол" on its own line, then double line break, then start your analysis.
  DO NOT redefine the animal traits. Focus PURELY on the friction or flow between the Elements. Explain who feeds whom, or who clashes with whom. Explain the emotional rhythm (e.g., who is rigid, who is fluid). Discuss health/physical compatibility (e.g., what organs or energies they support in each other). Conclude how they balance this dynamic.
  
  STYLE GUIDE REFERENCE (Model your structure, depth, and tone exactly after this):
  ${CONFIG.REFERENCES.PART_2}
  `;
  const r2 = callGeminiAPI(prompt2, apiKey, CONFIG.TEMPERATURE);

  // 3rd CALL: Mindset (Menge)
  const prompt3 = `
  ${SYSTEM_PROMPT}
  TASK: Write PART 3 ONLY (Mindset & Spiritual Connection).

  DATA:
  - Эрэгтэйн мэнгэ: ${male.mengeName} (${male.mengeElement} махбод)
  - Эмэгтэйн мэнгэ: ${female.mengeName} (${female.mengeElement} махбод)
  - Мэнгэний харилцаа: ${relations.menge}
  
  INSTRUCTIONS: Begin exactly with "3-р хэсэг. ОЮУН САНААНЫ ХОЛБООС: Мэнгэний Таарамж" on its own line, then double line break, then start your analysis.
  Analyze their intellectual compatibility and mindset based on Menge. How do their specific Menge elements (e.g., ${male.mengeElement} vs ${female.mengeElement}) interact intellectually? Explain the meaning of their relation ("${relations.menge}") in daily life. Does it cause arguments? How do they resolve them?
  
  STYLE GUIDE REFERENCE (Model your structure, depth, and tone exactly after this):
  ${CONFIG.REFERENCES.PART_3}
  `;
  const r3 = callGeminiAPI(prompt3, apiKey, CONFIG.TEMPERATURE);

  // 4th CALL: Lifestyle & Karma (Suudal)
  const prompt4 = `
  ${SYSTEM_PROMPT}
  TASK: Write PART 4 ONLY.
  
  DATA FOR PART 4:
  - Эрэгтэйн суудал: ${male.suudal} (${male.suudalElement} махбод)
  - Эмэгтэйн суудал: ${female.suudal} (${female.suudalElement} махбод)
  - Суудлын харилцаа: ${relations.suudal}
  - Шаардлагатай засал: ${relations.remedy.element} (${relations.remedy.advice})
  
  INSTRUCTIONS: Begin exactly with "4-р хэсэг. АМЬДРАЛЫН ХЭМНЭЛ: Суудлын Өнцөг & Гал Голомт" on its own line, then double line break, then start your analysis.
  Focus on their domestic life, finances, and home environment based on Suudal. Explain the Husband's Suudal lifestyle, then the Wife's. Explain their specific Suudal relation ("${relations.suudal}") - if it's bad (e.g., Хорт), explain *why* chemically (e.g., Fire burns Earth). Explain the practical remedy (${relations.remedy.advice}) using the 5 elements in their home.
  
  STYLE GUIDE REFERENCE (Model your structure, depth, and tone exactly after this):
  ${CONFIG.REFERENCES.PART_4}
  `;
  const r4 = callGeminiAPI(prompt4, apiKey, CONFIG.TEMPERATURE);

  // 5th CALL: Final Verdict
  const prompt5 = `
  ${SYSTEM_PROMPT}
  TASK: Write PART 5 ONLY.

  DATA FOR PART 5:
  - Нийт нийцлийн оноо: ${relations.score}%
  - Заслын зөвлөгөө: ${relations.remedy.advice}
  
  INSTRUCTIONS: Begin exactly with "5-р хэсэг. ҮЙЛИЙН ҮР БА ИРЭЭДҮЙН ЗУРАГЛАЛ: Нэгдсэн Дүгнэлт" on its own line, then double line break, then start your analysis.
  Give the final karmic verdict. Why did these two souls meet (e.g., a karmic debt, a healing bond)? Based on the ${relations.score}% score, give a realistic, honest conclusion. Emphasize that astrology is a guide, not a final sentence. End with a strong, beautiful blessing.
  
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

  const maxAttempts = 3;
  let lastErrorMsg = "";

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const res = UrlFetchApp.fetch(url, { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true });
    
    if (res.getResponseCode() === 429 || res.getResponseCode() >= 500) {
        lastErrorMsg = res.getContentText();
        Utilities.sleep(attempt * 2000); // 2s, then 4s wait
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
  
  const msg = CONFIG.UCHAT.DELIVERY_MESSAGE.replace("{{NAME}}", name);
  const btn = CONFIG.UCHAT.DELIVERY_BTN_TEXT;

  const payload = {
    user_ns: String(userNs).trim(), 
    data: {
      version: "v1",
      content: { messages: [ { type: "text", text: msg, buttons: [ { type: "url", caption: btn, url: pdfUrl } ] } ] }
    }
  };

  const maxAttempts = 3;
  let lastErrorMsg = "";

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const res = UrlFetchApp.fetch(CONFIG.UCHAT.ENDPOINT, {
      method: "post",
      headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
      payload: JSON.stringify(payload), muteHttpExceptions: true
    });

    const status = res.getResponseCode();
    const body = res.getContentText();

    if (status >= 200 && status < 300) {
      const json = JSON.parse(body);
      if (json.status === "ok" || json.success === true) {
        return; // Success
      } else {
        lastErrorMsg = "DELIVERY API failed: " + JSON.stringify(json);
      }
    } else {
      lastErrorMsg = "DELIVERY HTTP " + status + ": " + body.substring(0, 200);
    }

    // If not success, wait and retry
    Utilities.sleep(attempt * 2000);
  }

  throw new Error(`UChat Delivery Error after ${maxAttempts} attempts: ${lastErrorMsg}`);
}

// --- КОДЫН ТӨГСГӨЛ ---
