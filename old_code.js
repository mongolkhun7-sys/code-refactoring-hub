/**
 * PRODUCT: LOVE NUMEROLOGY REPORT GENERATOR
 * VERSION: v21.0 - ManyChat Only (Gold Ready)
 * AUTHOR: Jules + Codex
 * MODEL: gemini-2.5-flash
 ****************************************************************************************/

/*
MAINTAINER INSTRUCTIONS (FOR FUTURE THREADS / MODELS)
1) QUALITY RULE: CONFIG.REFERENCES.PART_1..PART_5 must contain FULL style anchor text blocks.
   Do NOT over-summarize references; style/persona drift starts from shortened references.
2) SECTION RULE: final report must include exact I, II, III, IV, V section headers.
3) UPDATE SCOPE (Product-level):
   - CONFIG.PRODUCT_NAME
   - CONFIG.REFERENCES.PART_1..PART_5
   - NUMEROLOGY mappings relevant to this product
   - CONFIG.DELIVERY_MESSAGE
4) DO-NOT-TOUCH (Core stability blocks unless explicitly requested):
   - main() status/retry flow
   - normalizeInputWithAI and validation gates
   - generateSequentialReport continuity/repair helpers
   - createPdf/processFormatting
5) DELIVERY RULE: this build is ManyChat-only. Column B must be ManyChat subscriber_id.
6) SOURCE FILES FOR NEXT THREAD:
   - /Users/munkhbat/Documents/zurhai-product-studio/gold-standard.txt
   - /Users/munkhbat/Documents/zurhai-product-studio/logic-spec.txt
*/

const CONFIG = {
  // --- SYSTEM CONFIGURATION ---
  VERSION: "v21.0-Love-Numerology-ManyChat",
  PRODUCT_NAME: "Хайр Дурлалын Нумерологи: Ирээдүйн Ханийн Код",
  SHEET_NAME: "Sheet1",
  BATCH_SIZE: 3,
  GEMINI_MODEL: "gemini-2.5-flash",
  TEMPERATURE: 0.35,
  MAX_RETRY_COUNT: 3,
  SECTION_MIN_CHARS: {
    I: 900,
    II: 1000,
    III: 900,
    IV: 850,
    V: 950
  },

  // --- COLUMN MAPPING (0-based Index) ---
  COLUMNS: {
    NAME: 0,      // A: User Name
    ID: 1,        // B: ManyChat subscriber_id
    INPUT: 2,     // C: Raw Input String
    PDF: 3,       // D: Output PDF URL
    STATUS: 4,    // E: Processing Status
    TOKEN: 5,     // F: Token Usage (Total)
    DEBUG: 6,     // G: Debug Data (JSON)
    DATE: 7,      // H: Timestamp
    VER: 8,       // I: Version Tag
    ERROR: 9      // J: Error Message
  },

  MAX_EXECUTION_TIME: 360000, // 6 Minutes
  SAFETY_BUFFER: 60000,       // 1 Minute buffer
  PROCESSING_STALE_MINUTES: 8,

  // ==================================================================================
  // 📜 GOLD STANDARD REFERENCE (LOVE NUMEROLOGY PRODUCT)
  // ==================================================================================
  REFERENCES: {
    PART_1: `
I ХЭСЭГ ХАЙР ДУРЛАЛЫН ДНХ
Хайрын Архетип
🔮 Таны төрсөн он сар өдрийн нийлбэр тооцооллоор 11 буюу 2 гэсэн Мастер тоо гарч байгаа нь таныг хайр дурлалын ертөнцөд Зөн совинтой Хайрлагч эсвэл Сэтгэлийн гүн холбоог эрхэмлэгч гэсэн маш онцгой мэдрэмтгий дүр төрхөөр тодорхойлж байна. Энэ нь таныг жирийн нэг гэрийн эзэгтэй эсвэл зүгээр л хүнтэй суух гэж яардаг эмэгтэй биш, харин эсрэг хүнийхээ дотоод ертөнц рүү өнгийж хардаг, сэтгэлийн нандин холбоог бүхнээс илүүд үздэг хүн гэдгийг илтгэдэг юм. Таны хайрлах арга барил бол нүдээрээ биш зүрхээрээ мэдрэх бөгөөд та үг хэлээгүй байхад л хамтрагчийнхаа сэтгэл санааг мэдэрч чаддаг байгалийн өгөгдөлтэй. Гэвч энэ өндөр мэдрэмж нь таныг харилцаанд хэт эмзэг болгодог бөгөөд хэрэв хамтрагч тань бүдүүлэг эсвэл сэтгэл хөдлөлгүй хүн байвал та маш хурдан шархалж, дотроосоо шаналж эхэлдэг тул та хайр дурлалаас зөвхөн халамж биш бас урам зориг, оюун санааны нэгдлийг хайж байдаг эмэгтэй юм.
Матрицын тоон өгөгдөл ба тайлал
Таны Пифагорын матрицын 9 нүдэнд 1111, 9999, 7 гэсэн онцгой тоонууд байрлаж байна. Эдгээр тоо тус бүр нь таны зан чанарт дараах эерэг болон сөрөг нөлөөллийг үзүүлдэг.
👑 Нэгдүгээрт, таны матрицад 1111 буюу нэгийн цифр дөрвөн удаа давтагдан байрлаж байгаа нь танд төрөлхийн маш хүчтэй Лидер зан чанарыг өгдөг. Энэ нь таныг бусдын эрхшээлд орох дургүй, бие даасан, хүчирхэг эмэгтэй болгодог сайн талтай ч, хайр дурлалын харилцаанд энэ хүч хэтэрвэл эрэгтэй хүнээ өөрийн мэдэлгүй захирах, дарах, эсвэл зөөлөн зан чанартай эрчүүдийг шүүмжилж голдог сүүдэр талыг үүсгэдэг тул та энэ хүчээ тэнцвэржүүлэх шаардлагатай байдаг.
🧠 Хоёрдугаарт, таны матрицад 9999 буюу есийн цифр дөрвөн удаа маш хүчтэй илэрч байгаа нь таныг "Гоц ухаантан"-ы түвшний оюуны чадамжтай, зөгнөлт мэт гайхалтай ой тогтоолттой болохыг харуулж байна. Та нэг уншсан, нэг сонссон зүйлээ хэзээ ч мартдаггүй. Энэ нь таныг маш ухаалаг хамтрагч болгодог ч, хайр дурлалын харилцаанд "Хэт их бодох" болон "Өнгөрснийг мартаж чадахгүй байх" асуудлыг хамгийн дээд цэгт нь хүргэдэг сөрөг талтай. Таны толгойд зогсоо зайгүй бодол эргэлдэж байдаг тул танд сэтгэлээ амрааж, тархиа "унтрааж" сурах нь аз жаргалтай байхын үндэс юм.
🛡️ Гуравдугаарт, таны матрицад 7-гийн тоо байрлаж байгаа нь танд "Тэнгэрийн хамгаалалт" буюу зөн совин, аливааг урьдчилан мэдрэх чадварыг өгдөг. Энэ нь таныг хүний мөн чанарыг зөв таних, аюулаас өөрийгөө хамгаалахад тусалдаг боловч, зарим тохиолдолд хэт их сэжиглэх эсвэл бодит байдлаас тасарч хэт төгс идеал ертөнцөд амьдрахад хүргэдэг талтай.
Хайрын Хэл
🗣️ Таны төрсөн өдөр 12 буюу 3-ын энергитэй байгаа нь таны Хайрын хэлийг Чин сэтгэлийн яриа болон Урмын үг гэж тодорхойлж байна. Та гаднаасаа 11-ийн тооны нөлөөгөөр нууцлаг, дуугүй юм шиг харагдаж болох ч яг хайртай хүнтэйгээ байхдаа өөрийгөө бүрэн илэрхийлэх, ярилцах, ойлголцох маш их хэрэгцээтэй байдаг. Танд зүгээр нэг үнэтэй бэлэг өгдөг залуугаас илүү тантай орой бүр цаг гарган ярилцдаг, таны мөрөөдлийг сонсдог мөн таныг байнга магтаж урамшуулдаг залуу хамгийн их таалагддаг. Хэрэв таны хамтрагч дуугүй дотогшоо эсвэл таныг сонсдоггүй хүн байвал та тэр харилцаанд ганцаардаж эхэлдэг бөгөөд энэ нь таны харилцаа хөрөх гол шалтгаан болдог тул таны хайрын түлхүүр бол Харилцан яриа бөгөөд таныг чин сэтгэлээсээ сонсдог хүн л таны зүрхийг эзэмшиж чадна.
`,
    PART_2: `
II ХЭСЭГ ЯАГААД ГАНЦ БИЕ БАЙНА ВЭ БА АЮУЛЫН ДОХИО
🚧 Кармын Өр
Таны төрсөн он сар өдрийн тооцооллыг нарийвчлан шалгаж үзэхэд 13, 14, 16, 19 гэсэн Кармын өрийн тусгай тоонууд илрээгүй нь таны хувьд маш том давуу тал бөгөөд үүнийг Нумерологид Цэвэр үйлийн үр буюу Цагаан цаас гэж нэрлэдэг юм. Энэ нь таныг өмнөх амьдралдаа бусдын хайр сэтгэлийг шархлуулсан эсвэл гэр бүлийн өмнө хүлээсэн үүргээсээ зугтсан ямар нэгэн хүнд ачаа тээшгүйгээр энэ ертөнцөд ирснийг баталж байгаа тул танд удам дамжсан ганц бие байдал эсвэл заяагдмал бүтэлгүйтэл гэж байхгүй гэсэн үг бөгөөд таны одоогийн ганц бие байдал нь зөвхөн таны өнөөдрийн сонголт болон шийдвэрээс шууд хамаарна.
🕳️ Матрицын Цоорхой
Таны матрицад 6 болон 8-ын тоо дутуу байгаа нь таны хайр дурлалын харилцаанд дараах хоёр гол гацааг үүсгэж байна. 
Нэгдүгээрт танд Өөрийн үнэ цэнэ болон Гэр бүлийн тав тухыг илэрхийлдэг 6-гийн тоо дутуу байгаа нь таныг хайр дурлалд орохоороо өөрийгөө бүрэн мартаж хамтрагчаасаа сэтгэл санааны хувьд хэт хамааралтай болох эсвэл өөрийгөө хайрлуулах эрхгүй мэтээр дорд үзэх эрсдэлтэй байдалд хүргэдэг. 
Хоёрдугаарт танд Хил хязгаар болон Бодит үр дүнг илэрхийлдэг 8-ын тоо дутуу байгаа нь таныг харилцаанд Үгүй гэж хэлж чадахгүй байх өөрийн орон зайг хамгаалж чадахгүй бусдад ашиглуулах эсвэл шийдэмгий алхам хийж чадахгүй эргэлзсээр цаг алдах сул талыг бий болгож байна.
🔒 Сэтгэл зүйн гацаа
Таны хайр дурлалын гол сэтгэл зүйн гацаа нь таны Мастер тоо 11-тэй холбоотой Төгс төгөлдөрийн урхи болон 9999 тооны Хэт гүнзгий ой тогтоолт хоёр нийлж үүсдэг. Та дотоод сэтгэлдээ ирээдүйн ханиа маш өндөр шалгуураар төсөөлдөгөөс гадна таны хүчирхэг 9999-ийн ой тогтоолт нь өнгөрсөн харилцааны шарх эсвэл хэзээ нэгэн цагт учирч байсан тэр онцгой нэгэнийг одоог хүртэл сэтгэлээсээ бүрэн тавьж явуулахгүй зууралдсаар байх магадлалтай. Та шинээр танилцсан хүн бүрийг тэрхүү өнгөрсний дурсамжтайгаа ухамсаргүйгээр харьцуулж тэр хүн шиг мэдрэмж өгөхгүй байна гэж голдог нь таны шинэ харилцаанд орох боломжийг хааж буй далд шалтгаан байж болзошгүй тул өнгөрснөөсөө бүрэн салах нь чухал юм.
⚠️ Аюулын бүс Энергийн үл зохицол ба сэрэмжлүүлэг
Таны өгөгдөл болох 1111 (мэдрэмтгий лидер), 9999 (хэт оюунлаг), 3 (харилцаа) гэсэн хослолд дараах гурван төрлийн зан төлөв болон энерги хамгийн хүчтэй сөрөг нөлөө үзүүлэх магадлалтай тул та ийм харилцаанаас сэрэмжлэх хэрэгтэй.
Нэгдүгээрт таны хувьд хэт захирангуй бүхнийг хяналтдаа байлгах гэсэн хандлагатай хүн хамгийн эрсдэлтэй. Та уг нь төрөлхийн хүчирхэг өгөгдөлтэй боловч ийм хүнтэй харилцах үедээ өөрийн үнэ цэнийг алдаж түүний хатуу дүрмэнд захирагдан өөрийнхөө мэдрэмж зөн совинг үгүйсгэхэд хүрдэг сөрөг талтай. Ийм хүн нь таныг Минийхээр бай гэж шахаж таныг өөрийн гэсэн орон зайгүй болгодог бөгөөд нумерологийн ухаанд энэ төрлийн хүчирхэг боловч дарангуйлах хандлагатай энерги нь ихэвчлэн сарын 8, 17, 26-ны өдөр төрсөн хүмүүст илүү тод илрэх магадлалтай байдаг тул та тэдний энэ талыг анхаарч харилцах нь зүйтэй.
Хоёрдугаарт сэтгэл зүйн хувьд тогтворгүй таашгүй зан авиртай хүн таны дотоод амар амгаланг хамгийн ихээр эвдэнэ. Таны 9999 код буюу зогсоо зайгүй ажилладаг тархи тань тодорхой бус байдлыг тэсвэрлэдэггүй тул нэг өдөр хайртай гэснээ маргааш нь алга болдог амлалт өгдөггүй хүнтэй байх нь таныг байнгын түгшүүр болон хэт их бодолд автуулж сэтгэл зүйн гүн ядаргаанд оруулдаг. Энэхүү эрх чөлөөг хэт эрхэмлэдэг тогтворгүй энерги нь ихэвчлэн сарын 5, 14, 23-ны өдөр төрсөн хүмүүст ажиглагддаг тул та ийм харилцаанд сэтгэл зүйгээ хамгаалах хэрэгтэй.
Гуравдугаарт сэтгэл хөдлөлөө илэрхийлдэггүй хэт дотогшоо эсвэл логик сэтгэлгээг шүтдэг хүн таныг сэтгэл санааны хувьд хамгийн их ганцаардуулна. Та бол ярилцаж ойлголцож байж хайрыг мэдэрдэг хүн боловч таны мэдрэмжийг "хэрэггүй зүйл" гэж шүүмжилдэг эсвэл ярилцах гэхээр чимээгүй болж сэтгэлийн хана босгодог хүнтэй амьдрах нь таныг хажуудаа хүнтэй мөртлөө яг л ганцаараа байгаа мэт мэдрэмжид унагаж аажмаар дотроос чинь унтраах аюултай. Энэхүү хөндий боловч гүн ухаанч энерги нь ихэвчлэн сарын 7, 16, 25-ны өдөр төрсөн хүмүүст давамгайлдаг тул та тэдэнтэй сэтгэл хөдлөлийн түвшинд ойлголцох эсэхээ сайтар нягтлах шаардлагатай.
`,
    PART_3: `
III ХЭСЭГ: ИРЭЭДҮЙН ХАНИЙН АВАТАР 
💎 Зан чанар
Таны заяаны ханийн дүр төрхийг Нумерологийн аргаар тооцоолж үзэхэд танд тохирох төгс зохицол нь таныг догдлуулж, савчуулдаг "Муу залуу" биш, харин таны дотоод шуургыг намжаадаг, сэтгэл зүйн хувьд маш тогтвортой "Дөлгөөн Халамжлагч" архетипийн хүн байх болно. Та Мастер 11 тоотой, дээр нь 1111 болон 9999 гэсэн маш өндөр ачаалалтай энергитэй хүн учраас танд гэртээ ирэхэд тань ямар ч шүүмжлэлгүйгээр сонсдог, таны хэт их бодлыг цэгцэлж өгдөг, маш өндөр сэтгэл хөдлөлийн оюун чадамжтай залуу таны жинхэнэ хань байхаар зурагдсан байна. Тэр хүн олны дунд хэт чанга дуугарч өөрийгөө дөвийлгөдөггүй, харин цаанаасаа л нэг тийм намуухан, уужуу тайван, бусдыг хүндэлдэг соёлтой зан чанартай байх бөгөөд түүний хажууд байхад танд өөрийнхөө тэр их хүчийг гаргах шаардлагагүй, яг л "Аюулгүй бүс"-д байгаа мэт тайван мэдрэмжийг төрүүлдэгээрээ бусдаас эрс ялгарна.
💼 Мэргэжил ба Статус
Мэргэжил болон нийгмийн статусын хувьд тэр хүн нь мөнгөний хойноос улайрч явдаг наймаачин эсвэл таныг захирах гэж оролддог бизнесмен гэхээсээ илүүтэй бусдын төлөө үйлчилдэг, оюуны хөдөлмөр эрхэлдэг эсвэл бүтээлч салбарын хүн байх магадлал маш өндөр байна. Тухайлбал эмч, багш, архитектор, сэтгэл зүйч, IT инженер эсвэл урлагийн мэдрэмжтэй бүтээлч захирал зэрэг мэргэжилтэй байж болох бөгөөд гол онцлог нь тэр хүн ажлаа зөвхөн мөнгө олох хэрэгсэл биш харин өөрийн үнэ цэнэ, бусдад туслах зорилго гэж хардаг "Зорилго төвтэй" хүн байх болно. Тэр хүн санхүүгийн хувьд тогтвортой байх боловч 8-ын тоотой хүмүүс шиг мөнгийг шүтэж, амьдралынхаа бүх цагийг ажилдаа зориулдаггүй, харин гэр бүл болон хайр сэтгэлдээ цаг гаргаж чаддаг тэнцвэртэй амьдралыг эрхэмлэдэг хүн байна.
🔢 Тохирох тоонууд
Таны энергитэй хамгийн төгс зохицох "Алтан тоонууд" буюу ирээдүйн ханийн тань төрсөн өдөр нь сарын 2, 6, 11, 20, 24, 29-ний өдрүүд байх магадлалтай. Эндээс хамгийн чухал нь 6-гийн тоотой (сарын 6, 15, 24-нд төрсөн) хүмүүс юм. Учир нь таны матрицад 6-гийн тоо дутуу байгааг бид өмнөх хэсэгт дурдсан. Тиймээс 6-гийн энергитэй залуус таны дутууг нөхөж, танд байхгүй тэрхүү гэр бүлсэг, халамжтай, тогтвортой байдлыг бэлэглэж чаддаг. Харин 2-ын тоотой (сарын 2, 11, 20, 29-нд төрсөн) хүмүүс нь тантай яг ижил сүнслэг түвшинд ойлголцож, таны 1111 болон 9999 кодыг үггүйгээр мэдэрч чадах Сүнсний ихэр тань байх болно.
✨ Гадаад төрхийн шинж
Гадаад төрхийн хувьд тэр хүн тийм ч хурц ширүүн харцтай биш, зөөлөн дулаан харцтай, цэвэрч нямбай хувцасладаг, магадгүй урлагийн мэдрэмжтэй, бага зэрэг романтик хүн байх бөгөөд та түүнийг анх харахад л "Бид өмнө нь уулзаж байсан билүү?" гэмээр маш дотно, танил мэдрэмж төрөх болно.
`,
    PART_4: `
IV ХЭСЭГ: ХАЙР ДУРЛАЛЫН 3 ЖИЛИЙН ЗУРАГЛАЛ 
2026 он: Эрх мэдэл ба Үр дүн
⏳ Таны 2026 оны "Хувийн жил"-ийн энергийг тооцоолж үзэхэд та энэ жил "Эрх мэдэл ба Үр дүн" буюу 8-ын тооны маш хүчирхэг мөчлөгт орж байна. Энэ нь таны хувьд хайр дурлалын түүхэндээ цоо шинэ хуудас нээх онцгой жил байх болно. 8-ын жил бол Кармын буцалт ирдэг буюу таны өнгөрсөн хугацаанд өөртөө оруулсан хөрөнгө оруулалт, тэвчээр хатуужлын шагнал бодитоор ирдэг жил юм. Хэрэв та дотоод сэтгэлээ цэгцэлж чадсан бол яг энэ жил таны амьдралын түвшинг дээшлүүлэх, нийгмийн статус өндөртэй, амжилттай яваа хүчирхэг хүнтэй учрах магадлал хамгийн өндөр байна. Одод эрхсийн хөдөлгөөнөөр таны "Соронзон таталцал" хамгийн оргилдоо хүрэх үеүүд нь таны төрсөн сар болох хаврын 4-р сар болон намрын эхэн үе буюу 9-р сарууд байх тул яг энэ үеүүдэд та гэртээ суухгүйгээр нийгмийн харилцаанд идэвхтэй оролцож, өөрийгөө нээлттэй байлгавал хувь заяаны гэнэтийн учрал таныг хүлээж байна.
2027 он: Төгсгөл ба Цэвэрлэгээ
🧹 Харин дараагийн жил буюу 2027 он нь Нумерологид "Төгсгөл ба Шилжилт" буюу 9-ийн жил тохиох бөгөөд энэ нь таны хайр дурлалын амьдралд "Их Цэвэрлэгээ" хийгдэх жил байх болно. Таны матрицад 9999 гэсэн хүчтэй ой тогтоолтын код байгаа тул та өнгөрснөөсөө салж чаддаггүй сул талтай гэдгийг бид дээр дурдсан. Тэгвэл 2027 он яг энэ асуудлыг шийдэж, өмнөх бүх эргэлзээ, айдас болон хуучин дурсамжуудаасаа бүрмөсөн салж, зөвхөн өөрийн сонгосон тэр хүнтэйгээ ирээдүйгээ холбох сэтгэл зүйн бэлтгэлээ хангах жил юм. Хэрэв 2026 онд учирсан харилцаа тань жинхэнэ хайр мөн бол энэ жил та хоёрын дунд үл ойлголцох бүх зүйлс шийдэгдэж, харилцаа тань цаг хугацааны шалгуурыг даван шалгарч үлдэх ба энэ жил бол хагацал салалт биш харин "Ганц бие амьдралаа үдэх" жил гэдгийг санаж өөрийгөө бүрэн чөлөөлөх хэрэгтэй.
2028 он: Шинэ эхлэл ба Бат бэх суурь
💒 Эцэст нь 2028 он бол таны хувьд цоо шинэ 9 жилийн мөчлөгийн эхлэл буюу "Хувийн жил 1" тохиож байгаа нь гэрлэлтээ албан ёсоор батлуулах, хурим найраа хийх, хүүхэдтэй болох эсвэл өөрийн гэсэн гэр бүлийн гал голомтоо тусдаа асаахад хамгийн ивээлтэй жил байх болно. Энэ онд таны амьдралд өрнөх үйл явдлууд ирээдүйн урт хугацааны жаргалтай амьдралын бат бөх суурь болох тул 2026 онд учирсан хүнтэйгээ 2027 онд итгэлцлээ бэхжүүлж, 2028 онд албан ёсоор нэг гэрт орох нь таны хайр дурлалын хамгийн зөв бөгөөд төгс сценари байхаар зурагдсан байна.
Учрах магадлалтай газрууд
📍 Таны Мастер 11 болон 7-гийн тооны мэдрэмтгий энерги нь хэт их чимээ шуугиантай, эмх замбараагүй газруудад хаагддаг тул та ирээдүйн ханиа баар цэнгээний газраас бус, харин өөрийнхөө сүнслэг болон оюунлаг мөн чанарт тохирсон орчин болох хувь хүний хөгжлийн сургалт, урлагийн үзэсгэлэн, номын нээлт эсвэл дээд зэрэглэлийн ажил хэргийн уулзалтууд дээрээс олох магадлал 90 хувьтай байна.
`,
    PART_5: `
V ХЭСЭГ: АМЖИЛТЫН НУУЦ ТҮЛХҮҮР
🔑 Стратеги: Тэнцвэрийг олох
Таны хайр дурлалын амьдралыг өөрчлөх хамгийн чухал стратеги бол "Бууж өгөх урлаг"-т суралцах явдал юм. Танд байгаа 1111 буюу дөрвөн нэгийн хүч нь таныг бүх зүйлийг хяналтдаа байлгах, удирдах, өөрөө шийдэх зуршилтай болгосон байдаг тул та хайр дурлалын харилцаанд орохоороо өөрийн мэдэлгүй "Эрэгтэй дүр"-д тоглож эхэлдэг. Тиймээс та гэрийнхээ босгоор давахдаа тэрхүү хүчирхэг Лидер хувцсаа гадаа үлдээж, зөвхөн эмэгтэйлэг, зөөлөн, хүлээн авагч чанараа ил гаргаж сурах нь таныг жинхэнэ халамжтай эртэй учрахад туслах болно. Мөн таны 9999 кодоос үүдэлтэй "Хэт их бодох" зуршлыг дарахын тулд бясалгал хийх, байгальд алхах эсвэл тархиа амраах дасгалуудыг тогтмол хийж, толгой доторх дуу хоолойгоо намжааж сурах нь таны сэтгэл зүйн амар амгалан байдлын үндэс юм.
🧘‍♀️ Өглөөний шидэт тарни (Affirmations)
Таны далд ухамсарт суусан "Би ганцаараа байх ёстой", "Намайг хэн ч ойлгохгүй", "Бүх эрчүүд ижилхэн" гэх мэт сөрөг итгэл үнэмшлийг эвдэхийн тулд та дараах шидэт үгсийг өглөө бүр толинд харж өөртөө хэлж хэвшээрэй.
"Би өнгөрсөн бүх гомдлоо тавьж явууллаа, би одоо шинэ хайрыг хүлээж авахад бэлэн байна."
"Би өөрийнхөө үнэ цэнийг мэдэрч байна, би хайрлагдах эрхтэй эмэгтэй."
"Би хүчирхэг байх албагүй, би халамжлуулах эрхтэй, би сул дорой байхыг өөртөө зөвшөөрч байна."
🎨 Аз дуудах өнгө ба чулуу
Таны Мастер 11 тооны мэдрэмтгий энергийг хамгаалж, зөн совинг тань улам хурцлахын тулд та "Мөнгөлөг цагаан", "Усан цэнхэр" болон "Нил ягаан" өнгийн хувцас эсвэл аксесуар хэрэглэх нь маш эерэг нөлөөтэй. Эдгээр өнгө нь таныг тайвшруулж, бусдын сөрөг хэл амнаас хамгаалдаг. Харин таны энергийг цэвэрлэж, хайр дурлалыг дуудах байгалийн чулуу бол "Сарны чулуу" (Moonstone) болон "Аметист" (Amethyst) юм. Сарны чулуу нь таны эмэгтэйлэг энергийг идэвхжүүлж, зөөлөн дулаан аураг үүсгэдэг бол Аметист чулуу нь таны 9999 кодоос үүдэлтэй стресс, бодлын урсгалыг намжааж, тайван нойрсоход тусална.
📜 Эцсийн дүгнэлт
Та бол жирийн нэг эмэгтэй биш, хорвоо ертөнцийг зүрхээрээ мэдэрдэг, оюун ухаанаараа нэвт хардаг онцгой өгөгдөлтэй "Мастер" хүн юм. Таны ганц бие байгаа шалтгаан нь таны муухай эсвэл азгүйд биш, харин таны шалгуур өндөр, таны сүнс зүгээр нэг энгийн харилцааг биш "Тэнгэрлэг холбоо"-г хайж байгаад оршино. Одоо та өөрийнхөө хүчийг (1111) зөөлөлж, ухаанаа (9999) амрааж, дутуугаа (6, 8) хүлээн зөвшөөрч чадвал таныг 2026 онд хувь заяаны гайхалтай бэлэг хүлээж байна.
`
  },

  // --- DELIVERY COPY ---
  DELIVERY_MESSAGE: `Сайн байна уу, {NAME}?

Таны "Хайр Дурлалын Нумерологи: Ирээдүйн Ханийн Код" бэлэн боллоо.

Файл: {URL}

(Татаж аваад хадгалаарай)`
};

const STATUS = {
  PROCESSING: "PROCESSING",
  DONE: "DONE",
  NEEDS_REVIEW: "NEEDS_REVIEW",
  RETRY_PREFIX: "RETRY_"
};

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
    1: [1, 10, 19, 28],
    2: [2, 11, 20, 29],
    3: [3, 12, 21, 30],
    4: [4, 13, 22, 31],
    5: [5, 14, 23],
    6: [6, 15, 24],
    7: [7, 16, 25],
    8: [8, 17, 26],
    9: [9, 18, 27],
    11: [1, 10, 11, 19, 20, 28, 29],
    22: [2, 4, 8, 11, 20],
    33: [3, 6, 9, 12, 21]
  },
  GENDER_STYLE_MAP: {
    "Эрэгтэй": {
      persona: "Эрэгтэй хэрэглэгчид зориулсан шууд, тод, хүндэтгэсэн хэллэг хэрэглэнэ.",
      preferred: "эрэгтэй хүн, та",
      forbidden: "эмэгтэй хүн, бүсгүй",
      replaceRules: [
        { from: "эмэгтэй хүн", to: "эрэгтэй хүн" },
        { from: "бүсгүй", to: "хүн" },
        { from: "охин минь", to: "хүн" }
      ]
    },
    "Эмэгтэй": {
      persona: "Эмэгтэй хэрэглэгчид зориулсан зөөлөн атлаа хүчтэй, ойлгомжтой хэллэг хэрэглэнэ.",
      preferred: "эмэгтэй хүн, та",
      forbidden: "эр хүн, ноён",
      replaceRules: [
        { from: "эр хүн", to: "эмэгтэй хүн" },
        { from: "ноён", to: "хүн" },
        { from: "хүү минь", to: "хүн" }
      ]
    }
  }
};

// --- MAIN FUNCTION ---
function main() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) throw new Error(`SYSTEM: "${CONFIG.SHEET_NAME}" sheet олдсонгүй.`);

    const rows = sheet.getDataRange().getValues();
    const KEYS = {
      GEMINI: PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY"),
      MANYCHAT: PropertiesService.getScriptProperties().getProperty("MANYCHAT_API_TOKEN"),
      TEMPLATE: PropertiesService.getScriptProperties().getProperty("TEMPLATE_ID")
    };
    validateKeys(KEYS);

    let processedCount = 0;
    const startTime = Date.now();

    for (let i = 1; i < rows.length; i++) {
      if (Date.now() - startTime > CONFIG.MAX_EXECUTION_TIME - CONFIG.SAFETY_BUFFER) break;
      if (processedCount >= CONFIG.BATCH_SIZE) break;

      const row = rows[i];
      const rowNumber = i + 1;
      const rawInput = String(row[CONFIG.COLUMNS.INPUT] || "").trim();
      const currentStatus = normalizeStatus(row[CONFIG.COLUMNS.STATUS]);
      const currentPdf = String(row[CONFIG.COLUMNS.PDF] || "").trim();
      const currentError = String(row[CONFIG.COLUMNS.ERROR] || "").trim();
      const lastUpdated = row[CONFIG.COLUMNS.DATE];

      if (!shouldProcessRow(rawInput, currentStatus, currentPdf, lastUpdated)) continue;

      sheet.getRange(rowNumber, CONFIG.COLUMNS.STATUS + 1).setValue(STATUS.PROCESSING);
      sheet.getRange(rowNumber, CONFIG.COLUMNS.DATE + 1).setValue(getNowStamp());
      SpreadsheetApp.flush();

      try {
        if (shouldRetryDeliveryOnly(currentStatus, currentPdf, currentError)) {
          const displayName = String(row[CONFIG.COLUMNS.NAME] || "").trim() || "Эрхэм";
          sendManyChat(row[CONFIG.COLUMNS.ID], currentPdf, displayName, KEYS.MANYCHAT);
          const now = getNowStamp();

          sheet.getRange(rowNumber, CONFIG.COLUMNS.STATUS + 1).setValue(STATUS.DONE);
          sheet.getRange(rowNumber, CONFIG.COLUMNS.DATE + 1).setValue(now);
          sheet.getRange(rowNumber, CONFIG.COLUMNS.VER + 1).setValue(CONFIG.VERSION);
          sheet.getRange(rowNumber, CONFIG.COLUMNS.ERROR + 1).setValue("");
          processedCount++;
          continue;
        }

        let pdfResult = null;
        let deliveryDone = false;
        try {
          const profile = parseAndCalculateProfile(rawInput, KEYS.GEMINI);
          const reportResult = generateSequentialReport(profile, KEYS.GEMINI);

          if (!reportResult || !reportResult.text || reportResult.text.length < 400) {
            throw new Error("GENERATE: Тайлан хангалтгүй богино үүссэн.");
          }

          pdfResult = createPdf(profile.name || "Эрхэм", reportResult.text, KEYS.TEMPLATE);
          sheet.getRange(rowNumber, CONFIG.COLUMNS.PDF + 1).setValue(pdfResult.url);
          sheet.getRange(rowNumber, CONFIG.COLUMNS.DATE + 1).setValue(getNowStamp());
          SpreadsheetApp.flush();

          sendManyChat(row[CONFIG.COLUMNS.ID], pdfResult.url, profile.name || "Эрхэм", KEYS.MANYCHAT);
          deliveryDone = true;

          const tokenBreakdown = buildTokenBreakdown(profile, reportResult);
          const totalTokens = tokenBreakdown.totalTokens || (profile.parsingUsage || 0) + (reportResult.usage || 0);
          const now = getNowStamp();

          sheet.getRange(rowNumber, CONFIG.COLUMNS.PDF + 1).setValue(pdfResult.url);
          sheet.getRange(rowNumber, CONFIG.COLUMNS.STATUS + 1).setValue(STATUS.DONE);
          sheet.getRange(rowNumber, CONFIG.COLUMNS.TOKEN + 1).setValue(totalTokens);
          sheet.getRange(rowNumber, CONFIG.COLUMNS.DEBUG + 1).setValue(JSON.stringify({
            profile: profile,
            token_breakdown: tokenBreakdown
          }));
          sheet.getRange(rowNumber, CONFIG.COLUMNS.DATE + 1).setValue(now);
          sheet.getRange(rowNumber, CONFIG.COLUMNS.VER + 1).setValue(CONFIG.VERSION);
          sheet.getRange(rowNumber, CONFIG.COLUMNS.ERROR + 1).setValue("");

          processedCount++;
        } catch (innerErr) {
          if (pdfResult && pdfResult.fileId && !deliveryDone && !isDeliveryError(innerErr)) {
            try {
              DriveApp.getFileById(pdfResult.fileId).setTrashed(true);
            } catch (cleanupErr) {
              console.warn("PDF cleanup warning:", cleanupErr);
            }
          }
          throw innerErr;
        }
      } catch (err) {
        console.error(err);
        applyRetryStatus(sheet, rowNumber, currentStatus, err);
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function validateKeys(keys) {
  if (!keys.GEMINI) throw new Error("SYSTEM: GEMINI_API_KEY байхгүй.");
  if (!keys.TEMPLATE) throw new Error("SYSTEM: TEMPLATE_ID байхгүй.");
  if (!keys.MANYCHAT) throw new Error("SYSTEM: MANYCHAT_API_TOKEN байхгүй.");
}

function getNowStamp() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
}

function createUsageStats() {
  return {
    totalTokens: 0,
    inputTokens: 0,
    outputTokens: 0,
    requests: 0
  };
}

function mergeUsageStats(target, detail) {
  if (!target || !detail) return target;
  target.totalTokens += Number(detail.totalTokens || 0);
  target.inputTokens += Number(detail.inputTokens || 0);
  target.outputTokens += Number(detail.outputTokens || 0);

  if (detail.requests !== undefined) {
    target.requests += Number(detail.requests || 0);
  } else if (
    Number(detail.totalTokens || 0) > 0 ||
    Number(detail.inputTokens || 0) > 0 ||
    Number(detail.outputTokens || 0) > 0
  ) {
    target.requests += 1;
  }
  return target;
}

function cloneUsageStats(stats) {
  return {
    totalTokens: Number((stats && stats.totalTokens) || 0),
    inputTokens: Number((stats && stats.inputTokens) || 0),
    outputTokens: Number((stats && stats.outputTokens) || 0),
    requests: Number((stats && stats.requests) || 0)
  };
}

function buildTokenBreakdown(profile, reportResult) {
  const parseStats = cloneUsageStats(profile && profile.parsingUsageDetail ? profile.parsingUsageDetail : createUsageStats());
  const reportStats = cloneUsageStats(reportResult && reportResult.usageDetail ? reportResult.usageDetail : createUsageStats());

  const total = Number(parseStats.totalTokens || 0) + Number(reportStats.totalTokens || 0);
  const input = Number(parseStats.inputTokens || 0) + Number(reportStats.inputTokens || 0);
  const output = Number(parseStats.outputTokens || 0) + Number(reportStats.outputTokens || 0);
  const requests = Number(parseStats.requests || 0) + Number(reportStats.requests || 0);
  const inputPct = total > 0 ? Number(((input / total) * 100).toFixed(1)) : 0;
  const outputPct = total > 0 ? Number(((output / total) * 100).toFixed(1)) : 0;

  return {
    totalTokens: total,
    inputTokens: input,
    outputTokens: output,
    inputPct: inputPct,
    outputPct: outputPct,
    requests: requests,
    parse: parseStats,
    report: reportStats
  };
}

function normalizeStatus(statusCell) {
  return String(statusCell || "").trim().toUpperCase();
}

function getRetryCount(status) {
  const m = String(status || "").match(/^RETRY_(\d+)$/i);
  return m ? Number(m[1]) : 0;
}

function shouldProcessRow(rawInput, status, pdfUrl, lastUpdated) {
  if (!rawInput) return false;
  if (pdfUrl && status !== STATUS.PROCESSING && !status.startsWith(STATUS.RETRY_PREFIX)) return false;
  if (status === STATUS.DONE || status === STATUS.NEEDS_REVIEW) return false;
  if (status.startsWith("ERROR_")) return false;
  if (status === STATUS.PROCESSING) return isStaleProcessing(lastUpdated);

  if (!status) return true;

  return status === "NEW" || status.startsWith(STATUS.RETRY_PREFIX);
}

function isStaleProcessing(lastUpdated) {
  const ts = parseSheetTimestamp(lastUpdated);
  if (!ts) return true;
  const diff = Date.now() - ts;
  return diff >= CONFIG.PROCESSING_STALE_MINUTES * 60 * 1000;
}

function parseSheetTimestamp(value) {
  if (!value) return null;

  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return value.getTime();
  }

  const s = String(value).trim();
  if (!s) return null;

  // Supports "yyyy-MM-dd HH:mm" (sheet stamp) and browser-native date strings.
  const normalized = s.includes("T") ? s : s.replace(" ", "T");
  const dt = new Date(normalized);
  return isNaN(dt.getTime()) ? null : dt.getTime();
}

function detectFailureType(err) {
  const msg = String(err && err.message ? err.message : err);
  if (msg.startsWith("NORMALIZE:")) return "ERROR_NORMALIZE";
  if (msg.startsWith("DELIVERY:")) return "ERROR_DELIVERY";
  if (msg.startsWith("PDF:")) return "ERROR_PDF";
  return "ERROR_GENERATE";
}

function isDeliveryError(err) {
  const msg = String(err && err.message ? err.message : err);
  return msg.startsWith("DELIVERY:");
}

function shouldRetryDeliveryOnly(status, pdfUrl, errorCell) {
  const normalizedStatus = String(status || "");
  const normalizedPdf = String(pdfUrl || "").trim();
  const normalizedError = String(errorCell || "").trim().toUpperCase();
  return (
    !!normalizedPdf &&
    normalizedStatus.startsWith(STATUS.RETRY_PREFIX) &&
    normalizedError.startsWith("ERROR_DELIVERY:")
  );
}

function applyRetryStatus(sheet, rowNumber, previousStatus, err) {
  const now = getNowStamp();
  const failureType = detectFailureType(err);
  const message = `${failureType}: ${String(err.message || err)}`;

  const currentRetry = getRetryCount(previousStatus);
  const nextRetry = currentRetry + 1;

  if (nextRetry <= CONFIG.MAX_RETRY_COUNT) {
    sheet.getRange(rowNumber, CONFIG.COLUMNS.STATUS + 1).setValue(`${STATUS.RETRY_PREFIX}${nextRetry}`);
  } else {
    sheet.getRange(rowNumber, CONFIG.COLUMNS.STATUS + 1).setValue(STATUS.NEEDS_REVIEW);
  }

  sheet.getRange(rowNumber, CONFIG.COLUMNS.ERROR + 1).setValue(message);
  sheet.getRange(rowNumber, CONFIG.COLUMNS.DATE + 1).setValue(now);
  sheet.getRange(rowNumber, CONFIG.COLUMNS.VER + 1).setValue(CONFIG.VERSION);
}

// ==========================================
// 1. CORE LOGIC ENGINE (AI POWERED PARSING)
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
  const genderStyle = getGenderStyle(normalized.gender);

  return {
    name: normalized.name,
    gender: normalized.gender,
    genderPersona: genderStyle.persona,
    genderPreferred: genderStyle.preferred,
    genderForbidden: genderStyle.forbidden,
    dob: normalized.date,
    age: age,
    parsingUsage: normalized.usage,
    parsingUsageDetail: normalized.usageDetail || createUsageStats(),

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

function buildMatrixSummary(counts) {
  const chunks = [];
  for (let i = 1; i <= 9; i++) {
    chunks.push(`${i}:${counts[i]}`);
  }
  return chunks.join(" | ");
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
  return {
    years: years,
    text: out.join("; ")
  };
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

function normalizeInputWithAI(raw, key) {
  const rawDate = normalizeDateValue(raw);
  const rawGender = detectGenderFromRaw(raw);
  if (rawDate && rawGender) {
    return {
      date: rawDate,
      gender: rawGender,
      name: normalizeName(raw, null),
      usage: 0,
      usageDetail: createUsageStats()
    };
  }

  const prompt = `
  TASK: Convert the input into strict JSON.
  INPUT: "${raw}"

  RULES:
  - Date must be YYYY.MM.DD
  - Gender must be "Эрэгтэй" or "Эмэгтэй"
  - Name optional, default "Эрхэм"
  - Return JSON only, no markdown

  JSON:
  {"date":"YYYY.MM.DD","gender":"Эрэгтэй","name":"Эрхэм"}
  `;

  let aiData = {};
  const usageStats = createUsageStats();

  try {
    const result = callGemini(prompt, key, 2, "NORMALIZE_INPUT");
    mergeUsageStats(usageStats, result.usageDetail);
    aiData = extractJsonBlock(result.text);
  } catch (e) {
    console.warn("Normalize AI warning:", e);
  }

  const aiDate = normalizeDateValue(aiData && aiData.date);
  if (rawDate && aiDate && rawDate !== aiDate) {
    console.warn(`Normalize mismatch: raw=${rawDate}, ai=${aiDate}, picked=raw`);
  }
  const date = rawDate || aiDate;
  if (!date) throw new Error("NORMALIZE: Төрсөн огноо зөв танигдсангүй.");

  const gender = normalizeGender(raw, aiData && aiData.gender);
  const name = normalizeName(raw, aiData && aiData.name);

  return { date, gender, name, usage: usageStats.totalTokens, usageDetail: usageStats };
}

// ==========================================
// 2. SEQUENTIAL GENERATION (BRIDGE ARCHITECTURE)
// ==========================================

function generateSequentialReport(data, apiKey) {
  const headers = {
    I: "**I ХЭСЭГ ХАЙР ДУРЛАЛЫН ДНХ**",
    II: "**II ХЭСЭГ ЯАГААД ГАНЦ БИЕ БАЙНА ВЭ БА АЮУЛЫН ДОХИО**",
    III: "**III ХЭСЭГ: ИРЭЭДҮЙН ХАНИЙН АВАТАР**",
    IV: "**IV ХЭСЭГ: ХАЙР ДУРЛАЛЫН 3 ЖИЛИЙН ЗУРАГЛАЛ**",
    V: "**V ХЭСЭГ: АМЖИЛТЫН НУУЦ ТҮЛХҮҮР**"
  };

  const hardFacts = buildHardFactsBlock(data);
  const requiredExcessNumbers = getRequiredExcessNumbers(data);
  const missingLock = buildMissingNumbersLock(data);

  const systemPrompt = `
ROLE: Хайр дурлалын нумерологич (Premium persona).
LANGUAGE: Natural Mongolian only.
STYLE: Native, confident, deep explanation. Энгийн ойлгомжтой боловч хэт товчлохгүй.
FORMAT: Bold ONLY for section headers.
RULE: Use exact user data and avoid contradictions.
GENDER RULE: Respect user's gender strictly. If wording is uncertain, use neutral "хүн".
OUTPUT RULE: Start immediately from the required section header.
FORBIDDEN OPENINGS: Never start with "За ойлголоо", "Ойлголоо", "Мэдлээ", "Тийм ээ", "OK".
FACT RULE: Never recalculate numbers/dates already provided in HARD FACTS.
MISSING RULE: For relationship diagnosis, prioritize only "Дутуу тоонууд (онцлох)".
COVERAGE RULE: In Section I, every number listed in "Илүүдэл тоонууд" must have explicit matrix interpretation.
NO LIST RULE: Do not use markdown bullets (*, -, •). Use cohesive paragraphs.
DEPTH RULE: Do not summarize into short generic text; develop full reasoning like reference style.
SUBHEADING RULE: Keep the expected subsection flow of each part.
REFERENCE RULE: REFERENCE STYLE text is persona/format anchor only. Never copy sample numbers or dates unless they match HARD FACTS.
`;

  const prompt1 = `
${systemPrompt}
TASK: Write Part 1 only.
Must include these exact headers:
${headers.I}
${headers.II}

USER DATA:
- Төрсөн огноо: ${data.dob} (${data.gender})
- Амьдралын зам: ${data.lifePath} (${data.lifePathName})
- Амьдралын замын тайлбар: ${data.lifePathDesc}
- Амьдралын замын тооцоо: ${data.lifePathCalcString}
- Өдрийн тоо: ${data.dayNumber}
- Матриц: ${data.matrixSummary}
- Илүүдэл тоонууд: ${JSON.stringify(data.excessNumbers)}
- Матрицад заавал тайлбарлах илүүдэл тоонууд: ${requiredExcessNumbers.join(", ") || "байхгүй"}
- Дутуу тоонууд (бүрэн): ${JSON.stringify(data.missingNumbers)}
- Кармын өр: ${data.karmicDebt}

HARD FACTS:
${hardFacts}

STRICT PART-1 COVERAGE:
- Section I матрицын тайлал дотор "Илүүдэл тоонууд"-д байгаа бүх тоонд тусдаа тайлбар бич.
- Дутуу тоог зөвхөн энэ жагсаалтаар дурдах: ${missingLock}

REFERENCE STYLE:
${CONFIG.REFERENCES.PART_1}
${CONFIG.REFERENCES.PART_2}
`;
  const usageStats = createUsageStats();
  let part1LocalFallbackUsed = false;
  let res1 = null;
  try {
    res1 = callGemini(prompt1, apiKey, 3, "REPORT_PART_1");
  } catch (e) {
    if (!isProhibitedContentError(e)) throw e;
    part1LocalFallbackUsed = true;
    console.warn("[REPORT_PART_1] PROHIBITED_CONTENT detected. Using deterministic local fallback.");
    res1 = {
      text: buildDeterministicPart12Fallback(data, headers),
      usage: 0,
      usageDetail: createUsageStats()
    };
  }
  mergeUsageStats(usageStats, res1.usageDetail);

  let part1 = sanitizeGeneratedPart(res1.text, {
    startHeader: headers.I,
    stopHeaders: [headers.III, headers.IV, headers.V]
  });
  part1 = enforceHeaderAtTop(part1, headers.I, /I\s*ХЭСЭГ\s*ХАЙР\s*ДУРЛАЛЫН\s*ДНХ/i);

  let secI = extractSectionBody(part1, headers.I, [headers.II, headers.III, headers.IV, headers.V]);
  let secII = extractSectionBody(part1, headers.II, [headers.III, headers.IV, headers.V]);
  if (!part1LocalFallbackUsed) {
    secI = repairSectionBody("I", secI, CONFIG.SECTION_MIN_CHARS.I, data, apiKey, systemPrompt, tailText(part1, 700), headers, usageStats);
    secII = repairSectionBody("II", secII, CONFIG.SECTION_MIN_CHARS.II, data, apiKey, systemPrompt, tailText(secI, 700), headers, usageStats);
  }
  assertSectionQuality("I", secI, CONFIG.SECTION_MIN_CHARS.I, data);
  assertSectionQuality("II", secII, CONFIG.SECTION_MIN_CHARS.II, data);
  part1 = composeSections([
    { header: headers.I, body: secI },
    { header: headers.II, body: secII }
  ]);

  const prompt2 = `
${systemPrompt}
TASK: Write Part 2 only.
Must include exact headers:
${headers.III}
${headers.IV}

CONTEXT: Part 1 already written. Do not repeat.
CONTINUITY: Match tone and rhythm from this excerpt, but do NOT repeat it:
${tailText(part1, 650)}

USER DATA:
- Хүйс: ${data.gender}
- Амьдралын зам: ${data.lifePath} (${data.lifePathName})
- Дутуу тоонууд (бүрэн): ${JSON.stringify(data.missingNumbers)}
- Тохирох өдрийн тоонууд: ${data.compatibleNumbers}
- Эрсдэлийн энергийн бүлгүүд: ${data.riskProfiles}
- 3 жилийн прогноз: ${data.forecastText}

HARD FACTS:
${hardFacts}

FACT LOCK:
- "Дутуу тоо" зөвхөн ${missingLock} байж болно.
- Энэ жагсаалтад байхгүй "дутуу X" хэллэг бүү ашигла.

REFERENCE STYLE:
${CONFIG.REFERENCES.PART_3}
${CONFIG.REFERENCES.PART_4}
`;
  const res2 = callGemini(prompt2, apiKey, 3, "REPORT_PART_2");
  mergeUsageStats(usageStats, res2.usageDetail);

  let part2 = sanitizeGeneratedPart(res2.text, {
    startHeader: headers.III,
    stopHeaders: [headers.V]
  });
  part2 = removeLeadingOverlap(part1, part2);
  part2 = enforceHeaderAtTop(part2, headers.III, /III\s*ХЭСЭГ:\s*ИРЭЭДҮЙН\s*ХАНИЙН\s*АВАТАР/i);

  let secIII = extractSectionBody(part2, headers.III, [headers.IV, headers.V]);
  let secIV = extractSectionBody(part2, headers.IV, [headers.V]);
  secIII = repairSectionBody("III", secIII, CONFIG.SECTION_MIN_CHARS.III, data, apiKey, systemPrompt, tailText(part1, 700), headers, usageStats);
  secIV = repairSectionBody("IV", secIV, CONFIG.SECTION_MIN_CHARS.IV, data, apiKey, systemPrompt, tailText(secIII, 700), headers, usageStats);
  assertSectionQuality("III", secIII, CONFIG.SECTION_MIN_CHARS.III, data);
  assertSectionQuality("IV", secIV, CONFIG.SECTION_MIN_CHARS.IV, data);
  part2 = composeSections([
    { header: headers.III, body: secIII },
    { header: headers.IV, body: secIV }
  ]);

  const prompt3 = `
${systemPrompt}
TASK: Write Part 3 only.
Must include exact header:
${headers.V}

CONTEXT: Parts 1,2 already written. Do not repeat.
CONTINUITY: Match tone and rhythm from this excerpt, but do NOT repeat it:
${tailText(part2, 650)}

USER DATA:
- Хүйс: ${data.gender}
- Дутуу тоонууд (бүрэн): ${JSON.stringify(data.missingNumbers)}
- Амжилтын өнгө: ${data.luckyColor}
- Азын чулуу: ${data.luckyStone}
- Азын тоо: ${data.luckyNumbers}

HARD FACTS:
${hardFacts}

FACT LOCK:
- "Дутуу тоо" зөвхөн ${missingLock} байж болно.
- Энэ жагсаалтад байхгүй "дутуу X" хэллэг бүү ашигла.

REFERENCE STYLE:
${CONFIG.REFERENCES.PART_5}
`;
  const res3 = callGemini(prompt3, apiKey, 3, "REPORT_PART_3");
  mergeUsageStats(usageStats, res3.usageDetail);

  let part3 = sanitizeGeneratedPart(res3.text, {
    startHeader: headers.V,
    stopHeaders: []
  });
  part3 = removeLeadingOverlap(`${part1}

${part2}`, part3);
  part3 = enforceHeaderAtTop(part3, headers.V, /V\s*ХЭСЭГ:\s*АМЖИЛТЫН\s*НУУЦ\s*ТҮЛХҮҮР/i);

  let secV = extractSectionBody(part3, headers.V, []);
  secV = repairSectionBody("V", secV, CONFIG.SECTION_MIN_CHARS.V, data, apiKey, systemPrompt, tailText(part2, 700), headers, usageStats);
  assertSectionQuality("V", secV, CONFIG.SECTION_MIN_CHARS.V, data);
  part3 = composeSections([{ header: headers.V, body: secV }]);

  const rawReport = `${part1}

${part2}

${part3}`;
  const guardedReport = applyGenderLanguageGuard(rawReport, data.gender);
  const factGuardedReport = applyDeterministicFactGuards(guardedReport, data);
  const finalText = normalizeReportArtifacts(factGuardedReport);
  return {
    text: finalText,
    usage: usageStats.totalTokens,
    usageDetail: usageStats
  };
}

function isProhibitedContentError(err) {
  const msg = String(err && err.message ? err.message : err);
  return (
    msg.indexOf("promptBlock=PROHIBITED_CONTENT") !== -1 ||
    msg.indexOf("finishReason=SAFETY") !== -1
  );
}

function buildDeterministicPart12Fallback(data, headers) {
  const excessNums = getRequiredExcessNumbers(data);
  const matrixCounts = data.matrixCounts || {};
  const missingNums = (data.missingNumbers || []).map(m => Number(m.number)).filter(n => n >= 1 && n <= 9);

  const excessNarrative = excessNums.length
    ? excessNums.map(n => {
        const count = Number(matrixCounts[n] || 0);
        const info = NUMEROLOGY.MATRIX_EXCESS_MAP[n] || { title: "Энергийн төвлөрөл", desc: "Хэт туйлшрахаас сэргийлж тэнцвэр хадгална." };
        return `${n}-ийн тоо матрицад ${count} удаа давтагдсан нь ${info.title} хэлбэрээр илэрч, ${info.desc} Энэ ${n}-ийн энергийг харилцаанд хүндлэл, уян хатан байдалтайгаар ашиглах нь тогтвортой холбоонд хамгийн чухал.`;
      }).join(" ")
    : "Матрицын тоон өгөгдөл ба тайлалын хүрээнд илүүдэл тоо онцгойрон харагдахгүй байна. Ийм үед харилцааны тэнцвэрийг өдөр тутмын харилцаа, амлалт, бодит үйлдлээр тогтвортой барих нь хамгийн үр дүнтэй байдаг.";

  const missingNarrative = missingNums.length
    ? missingNums.map(n => {
        const info = NUMEROLOGY.MISSING_NUMBER_MAP[n] || { title: "Цоорхой", risk: "Харилцаанд тодорхойгүй байдал үүсч болзошгүй." };
        return `${n}-ийн тоо дутуу байгаа нь ${info.title} байдлаар илэрч, ${info.risk} Иймд энэ хэсгийг нөхөх дадал бол хил хязгаар, үнэнч харилцаа, тогтмол анхаарал юм.`;
      }).join(" ")
    : "Матрицын Цоорхой томоохон хэмжээнд ажиглагдахгүй байгаа тул гол анхаарал нь харилцааны тогтвортой хэмнэл, харилцан ойлголцол, хэлсэндээ хүрэх итгэлцлийг хадгалах талд төвлөрнө.";

  const sectionI = normalizeReportArtifacts(`
Хайрын Архетип
Таны төрсөн огноо ${data.dob}, хүйс ${data.gender}, амьдралын зам ${data.lifePath} (${data.lifePathName}) гэсэн гол өгөгдлөөс харахад та харилцаанд сэтгэлийн чанар, үнэнч байдал, дотоод аюулгүй мэдрэмжийг хамгийн түрүүнд эрэмбэлдэг хүн байна. Энэ архетип нь гял цал гаднах дүр төрхөөс илүү тогтвортой зан чанар, хариуцлагатай харилцаа, урт хугацааны зорилготой түншлэлийг илүүд үздэгээрээ онцлог. Таны хувьд хурдан сэтгэл хөдлөлөөс илүү бодит үйлдэл, амлалтаа биелүүлэх чадвар, тэнцвэртэй ярилцлага хамгийн өндөр үнэ цэнтэй тул харилцааны чанар шууд өсөх боломжтой.

Матрицын тоон өгөгдөл ба тайлал
Матрицын дүн ${data.matrixSummary} гэж гарч байгаа бөгөөд энэ нь таны зан төлөвт тодорхой давтамжууд хүчтэй ажиллаж байгааг харуулж байна. ${excessNarrative} Үүн дээр амьдралын замын тайлбар болох "${data.lifePathDesc}" чанарыг давхар харвал танд өөрийн хүчтэй талуудаа хэтрүүлэхгүй, нөгөө хүний орон зайг хүндэтгэх ухаалаг тэнцвэр хамгийн чухал байна.

Хайрын Хэл
Өдрийн тоо ${data.dayNumber} болон амьдралын замын нөлөөг нэгтгэн харахад таны хайрын хэл нь зөвхөн үгээр биш, тогтмол анхаарал, тайван ярилцлага, итгэл даасан үйлдлээр илэрдэг хэв маягтай. Та ойлголцлын хурдыг албадах бус, харин нээлттэй асуулт, сонсох чадвар, хүндэтгэлтэй хариу үйлдлээр харилцаагаа бататгах үед хамгийн сайн үр дүн гаргана. Энэ зарчмыг тогтмол баримтлах нь богино хугацааны үл ойлголцлыг багасгаж, урт хугацааны тогтвортой холбоог бий болгоно.
`);

  const sectionII = normalizeReportArtifacts(`
Кармын Өр
Кармын өрийн дүн: ${data.karmicDebt}. Энэ үзүүлэлтээс шалтгаалсан заавал давтагдах сөрөг сценари гэж шууд дүгнэхээс илүүтэй, өнөөдрийн сонголт, сэтгэлзүйн хариу үйлдэл, харилцааны дадал хэвшил бодит үр дүнг тодорхойлдог. Иймд өнгөрснийг тайлбарлахдаа хатуу онош тавихгүй, харин одоогийн шийдвэрийн чанарыг сайжруулах зарчим баримтлах нь зөв.

Матрицын Цоорхой
Матрицын дутуу тоонуудын үнэлгээ: ${missingNums.length ? missingNums.join(", ") : "байхгүй"}. ${missingNarrative} Энэ хэсэг нь айдас төрүүлэх дохио биш, харин хөгжүүлэх ур чадварын зураглал гэж ойлговол илүү үр дүнтэй бөгөөд өдөр тутамд жижиг боловч тогтмол алхмаар нөхөгдөх боломжтой.

Сэтгэл зүйн гацаа
Харилцаанд гацаа үүсдэг гол шалтгаан нь ихэнхдээ дутуу мэдээлэлтэй таамаг, хэлээгүй хүлээлт, эсвэл өөрийгөө хамгаалах автомат урвал байдаг. Таны хувьд хамгийн ажилладаг арга нь хэт ерөнхий дүгнэлтээс зайлсхийж, нэг асуудлыг нэг удаа, нэг зорилгоор ярилцах бүтэцтэй харилцаа юм. Мөн өөрийн хэрэгцээг товч тодорхой хэлж сурвал үл ойлголцол хурдан буурч, хоёуланд нь аюулгүй орон зай үүсдэг.

Аюулын бүс Энергийн үл зохицол ба сэрэмжлүүлэг
Эрсдэлийн энергийн бүлгүүд: ${data.riskProfiles}. Энэ мэдээллийг хүн шүүх хэрэгсэл биш, харин харилцааны эрт дохиог таних чиглүүлэгч гэж ашиглах нь зөв. Хэрэв хэт хяналт, тогтворгүй харилцааны мөчлөг, эсвэл сэтгэлээ хаах хандлага давтагдвал шууд шошголохын оронд хил хязгаараа тодорхой хэлж, бодит зан үйлийн өөрчлөлт ажиглагдах эсэхээр шийдвэр гаргаарай. Ингэснээр та сэтгэл хөдлөлөөр биш бодит нотолгоонд тулгуурласан аюулгүй сонголт хийх боломжтой.
`);

  return composeSections([
    { header: headers.I, body: sectionI },
    { header: headers.II, body: sectionII }
  ]);
}

function tailText(text, maxChars) {
  const s = String(text || "");
  if (!s) return "";
  return s.length <= maxChars ? s : s.slice(s.length - maxChars);
}

function sanitizeGeneratedPart(text, opts) {
  let out = normalizeFinalReport(text);
  out = stripConversationalLead(out);
  if (opts && opts.startHeader) out = trimBeforeHeader(out, opts.startHeader);
  if (opts && opts.stopHeaders && opts.stopHeaders.length) out = trimAfterHeaders(out, opts.stopHeaders);
  out = stripConversationalLead(out);
  return normalizeFinalReport(out);
}

function stripConversationalLead(text) {
  const leadPattern = /^(за+|за ойлголоо|ойлголоо|мэдлээ|тийм ээ|мэдээж|за тэгье|ok|okay)\b/i;
  const lines = String(text || "").split(/\n/);
  let cut = 0;
  while (cut < Math.min(lines.length, 6)) {
    const line = lines[cut].trim().replace(/^[-*•\d\.\)\s]+/, "");
    if (!line) {
      cut++;
      continue;
    }
    if (leadPattern.test(line)) {
      cut++;
      continue;
    }
    break;
  }
  return lines.slice(cut).join("\n").trim();
}

function trimBeforeHeader(text, header) {
  const s = String(text || "");
  const idx = s.indexOf(header);
  if (idx > -1) return s.slice(idx).trim();
  return s.trim();
}

function trimAfterHeaders(text, headers) {
  const s = String(text || "");
  let cutIndex = -1;
  for (let i = 0; i < headers.length; i++) {
    const idx = s.indexOf(headers[i]);
    if (idx > 0 && (cutIndex === -1 || idx < cutIndex)) {
      cutIndex = idx;
    }
  }
  if (cutIndex > 0) return s.slice(0, cutIndex).trim();
  return s.trim();
}

function removeLeadingOverlap(previousText, currentText) {
  const prev = String(previousText || "");
  let cur = String(currentText || "").trim();
  if (!prev || !cur) return cur;

  for (let len = Math.min(260, cur.length); len >= 80; len -= 10) {
    const probe = cur.slice(0, len).trim();
    if (probe.length < 80) continue;
    if (prev.indexOf(probe) !== -1) {
      const stripped = cur.slice(len).trim();
      if (stripped.length >= 120) return stripped;
    }
  }

  const firstBlock = (cur.split(/\n\s*\n/)[0] || "").trim();
  if (firstBlock.length >= 80 && prev.indexOf(firstBlock) !== -1) {
    const stripped = cur.slice(firstBlock.length).trim();
    if (stripped.length >= 120) cur = stripped;
  }

  return cur;
}

function hasHeaderRegex(text, pattern) {
  return pattern.test(String(text || ""));
}

function enforceHeaderAtTop(text, exactHeader, headerRegex) {
  let out = normalizeFinalReport(text);
  if (!hasHeaderRegex(out, headerRegex)) {
    return normalizeFinalReport(`${exactHeader}\n${out}`);
  }
  return trimBeforeHeader(out, exactHeader);
}

function removeKnownHeaders(text, headers) {
  let out = String(text || "");
  const list = [headers.I, headers.II, headers.III, headers.IV, headers.V];
  for (let i = 0; i < list.length; i++) {
    const h = list[i];
    if (!h) continue;
    const escaped = escapeRegExp(h);
    out = out.replace(new RegExp(escaped, "g"), "");
  }
  return normalizeFinalReport(out);
}

function extractSectionBody(text, startHeader, stopHeaders) {
  const src = normalizeFinalReport(text);
  const start = src.indexOf(startHeader);
  if (start === -1) return "";

  const from = start + startHeader.length;
  let to = src.length;
  const stops = stopHeaders || [];
  for (let i = 0; i < stops.length; i++) {
    const idx = src.indexOf(stops[i], from);
    if (idx > -1 && idx < to) to = idx;
  }
  return normalizeFinalReport(src.slice(from, to));
}

function isWeakSection(body, minChars) {
  const clean = normalizeFinalReport(body);
  if (!clean) return true;
  const paragraphCount = clean.split(/\n{2,}/).filter(Boolean).length;
  return clean.length < minChars || paragraphCount < 2;
}

function hasUnfinishedEnding(text) {
  const clean = normalizeFinalReport(text);
  if (!clean) return true;
  const lines = clean.split("\n").map(line => line.trim()).filter(Boolean);
  if (!lines.length) return true;
  const lastLine = lines[lines.length - 1];
  return !/[.!?…]["')»”]?$/.test(lastLine);
}

function getRequiredExcessNumbers(data) {
  const list = (data && data.excessNumbers ? data.excessNumbers : [])
    .map(item => Number(item && item.number))
    .filter(n => Number.isFinite(n) && n >= 1 && n <= 9);
  const uniq = [];
  for (let i = 0; i < list.length; i++) {
    if (uniq.indexOf(list[i]) === -1) uniq.push(list[i]);
  }
  return uniq.sort((a, b) => a - b);
}

function hasMatrixNumberNarrative(text, number) {
  const n = Number(number);
  if (!(n >= 1 && n <= 9)) return true;
  const src = String(text || "");
  if (!src.trim()) return false;

  const numberOrCode = `(?:${n}|${n}{2,5})`;
  const nounPattern = new RegExp(
    String.raw`(?:^|[^\d])${numberOrCode}(?:\s*[-–—]?\s*(?:ын|ийн|гийн)\s*(?:тоо|цифр|энерги)|\s*(?:тоо|цифр|энерги))(?:[^\d]|$)`,
    "i"
  );
  if (nounPattern.test(src)) return true;

  const matrixPattern = new RegExp(
    String.raw`матриц[^\n]{0,200}(?:^|[^\d])${numberOrCode}(?:[^\d]|$)`,
    "i"
  );
  if (matrixPattern.test(src)) return true;

  return false;
}

function containsPhraseLoose(text, phrase) {
  const src = String(text || "");
  const target = String(phrase || "").trim();
  if (!target) return true;
  const pattern = escapeRegExp(target).replace(/\s+/g, "\\s+");
  return new RegExp(pattern, "i").test(src);
}

function getSectionRequiredPhrases(sectionKey) {
  if (sectionKey === "I") {
    return ["Хайрын Архетип", "Матрицын тоон өгөгдөл ба тайлал", "Хайрын Хэл"];
  }
  if (sectionKey === "II") {
    return ["Кармын Өр", "Матрицын Цоорхой", "Сэтгэл зүйн гацаа", "Аюулын бүс Энергийн үл зохицол ба сэрэмжлүүлэг"];
  }
  if (sectionKey === "III") {
    return ["Зан чанар", "Мэргэжил ба Статус", "Тохирох тоонууд", "Гадаад төрхийн шинж"];
  }
  if (sectionKey === "IV") {
    return ["2026 он", "2027 он", "2028 он", "Учрах магадлалтай газрууд"];
  }
  if (sectionKey === "V") {
    return ["Стратеги: Тэнцвэрийг олох", "Өглөөний шидэт тарни", "Аз дуудах өнгө ба чулуу", "Эцсийн дүгнэлт"];
  }
  return [];
}

function getSectionCoverageIssues(sectionKey, body, data) {
  const missing = [];
  const requiredPhrases = getSectionRequiredPhrases(sectionKey);
  for (let i = 0; i < requiredPhrases.length; i++) {
    const phrase = requiredPhrases[i];
    if (!containsPhraseLoose(body, phrase)) {
      missing.push(`"${phrase}" дэд гарчиг`);
    }
  }

  if (sectionKey !== "I") return missing;

  const requiredExcess = getRequiredExcessNumbers(data);
  for (let i = 0; i < requiredExcess.length; i++) {
    const n = requiredExcess[i];
    if (!hasMatrixNumberNarrative(body, n)) {
      missing.push(`матрицын ${n}-ын тайлбар`);
    }
  }
  return missing;
}

function buildSectionIssueHint(sectionKey, issues) {
  if (!issues) return "";
  const lines = [];
  if (issues.coverage && issues.coverageDetails && issues.coverageDetails.length) {
    lines.push(`- Дутуу заавал орох тайлбар: ${issues.coverageDetails.join(", ")}`);
    lines.push(`- Section бүтэц, дэд гарчиг, тайлбарын бүрэн бүтэн байдлыг хадгал.`);
  }
  if (issues.unfinished) {
    lines.push(`- Төгсгөл тасархай биш, бүтэн өгүүлбэрээр төгсгө.`);
  }
  if (issues.weak) {
    lines.push(`- Хэт товч биш, дор хаяж хэд хэдэн бүтэн догол мөртэй гүн тайлбар өг.`);
  }
  return lines.join("\n");
}

function getSectionIssues(sectionKey, body, minChars, data) {
  const coverageDetails = getSectionCoverageIssues(sectionKey, body, data);
  return {
    weak: isWeakSection(body, minChars),
    unfinished: hasUnfinishedEnding(body),
    coverage: coverageDetails.length > 0,
    coverageDetails: coverageDetails
  };
}

function repairSectionBody(sectionKey, currentBody, minChars, data, apiKey, systemPrompt, continuity, headers, usageStats) {
  let body = normalizeFinalReport(currentBody);
  const maxAttempts = 2;
  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    const issues = getSectionIssues(sectionKey, body, minChars, data);
    if (!(issues.weak || issues.unfinished || issues.coverage)) break;
    let regen = "";
    try {
      regen = generateSectionBody(
        sectionKey,
        data,
        apiKey,
        systemPrompt,
        continuity,
        headers,
        usageStats,
        buildSectionIssueHint(sectionKey, issues)
      );
    } catch (e) {
      console.warn(`Section regenerate warning [${sectionKey}] attempt ${attempt + 1}:`, e);
      regen = "";
    }
    if (!regen) break;
    body = normalizeFinalReport(regen);
  }
  return normalizeReportArtifacts(body);
}

function assertSectionQuality(sectionKey, body, minChars, data) {
  const issues = getSectionIssues(sectionKey, body, minChars, data);
  if (issues.coverage) {
    throw new Error(`GENERATE: ${sectionKey} хэсэгт заавал орох тайлбар дутуу байна (${issues.coverageDetails.join(", ")}).`);
  }
  // Hard-fail only for real structural/content quality issues.
  if (issues.weak || issues.unfinished) {
    throw new Error(`GENERATE: ${sectionKey} хэсэг дутуу эсвэл тасархай байна.`);
  }
}

function composeSections(sections) {
  return normalizeFinalReport(
    (sections || [])
      .map(s => `${s.header}\n${normalizeFinalReport(s.body || "")}`)
      .join("\n\n")
  );
}

function ensureLeadLine(text, leadLine) {
  const body = normalizeFinalReport(text);
  const lead = normalizeFinalReport(leadLine);
  if (!body) return lead;
  if (body.indexOf(lead) === 0) return body;
  return `${lead}\n\n${body}`;
}
function buildHardFactsBlock(data) {
  const missingNums = (data.missingNumbers || []).map(m => m.number).join(", ") || "байхгүй";
  const excessNums = (data.excessNumbers || []).map(e => `${e.number} x${e.count}`).join(", ") || "байхгүй";
  const forecastPairs = (data.forecastYears || [])
    .map(y => `${y.year} -> Хувийн жил ${y.number} (${y.title})`)
    .join("; ");

  return [
    `Төрсөн огноо: ${data.dob}`,
    `Хүйс: ${data.gender}`,
    `Нас: ${data.age}`,
    `Амьдралын зам: ${data.lifePath} (${data.lifePathName})`,
    `Амьдралын замын тооцоо: ${data.lifePathCalcString}`,
    `Өдрийн тоо: ${data.dayNumber}`,
    `Матриц: ${data.matrixSummary}`,
    `Матрицын давтамж: ${JSON.stringify(data.matrixCounts || {})}`,
    `Илүүдэл тоо: ${excessNums}`,
    `Дутуу тоо: ${missingNums}`,
    `Кармын өр: ${data.karmicDebt}`,
    `Хосын зохицох өдрийн тоонууд: ${data.compatibleNumbers}`,
    `Эрсдэлийн бүлгүүд: ${data.riskProfiles}`,
    `Форекаст: ${forecastPairs}`,
    `Амжилтын өнгө: ${data.luckyColor}`,
    `Азын чулуу: ${data.luckyStone}`,
    `Азын тоо: ${data.luckyNumbers}`
  ].join("\n");
}

function normalizeMissingNumberList(missingNumbers) {
  const raw = missingNumbers || [];
  const nums = [];
  for (let i = 0; i < raw.length; i++) {
    const n = Number(raw[i] && raw[i].number);
    if (!(n >= 1 && n <= 9)) continue;
    if (nums.indexOf(n) === -1) nums.push(n);
  }
  return nums.sort((a, b) => a - b);
}

function buildMissingNumbersLock(data) {
  const nums = normalizeMissingNumberList(data && data.missingNumbers);
  return nums.length ? nums.join(", ") : "байхгүй";
}

function enforceLifePathFact(text, lifePathValue) {
  const lifePath = Number(lifePathValue);
  if (!(lifePath >= 1 && lifePath <= 33)) return String(text || "");
  return String(text || "").replace(/(амьдралын зам\s*)(\d{1,2})/gi, function(match, prefix, got) {
    return Number(got) === lifePath ? match : `${prefix}${lifePath}`;
  });
}

function enforceKarmicDebtFact(text, karmicDebtValue) {
  const debt = String(karmicDebtValue || "").trim();
  if (!debt) return String(text || "");
  return String(text || "").replace(/(Кармын өрийн дүн:\s*)([^\.\n]+)/gi, function(match, prefix, got) {
    return got.trim() === debt ? match : `${prefix}${debt}`;
  });
}

function enforceMissingNumberFacts(text, missingNumbers) {
  let out = String(text || "");
  const nums = normalizeMissingNumberList(missingNumbers);
  const canonical = nums.length ? nums.join(", ") : "байхгүй";

  out = out.replace(/(Матрицын дутуу тоонуудын үнэлгээ:\s*)([^\.\n]+)/gi, function(match, prefix, got) {
    return got.trim() === canonical ? match : `${prefix}${canonical}`;
  });

  if (nums.length !== 1) return out;

  const primary = String(nums[0]);
  const replaceIfWrong = function(match, prefix, got, suffix) {
    return Number(got) === Number(primary) ? match : `${prefix}${primary}${suffix}`;
  };

  out = out.replace(/(\bдутуу\s+байгаа\s+)([1-9])(-?гийн)/giu, replaceIfWrong);
  out = out.replace(/(\bдутуу\s+)([1-9])(-?гийн)/giu, replaceIfWrong);
  out = out.replace(/(\(дутуу\s*)([1-9])(\))/giu, replaceIfWrong);
  out = out.replace(/(\bдутуу\s+)([1-9])(\b)/giu, replaceIfWrong);

  return out;
}

function applyDeterministicFactGuards(text, data) {
  let out = String(text || "");
  out = enforceLifePathFact(out, data && data.lifePath);
  out = enforceKarmicDebtFact(out, data && data.karmicDebt);
  out = enforceMissingNumberFacts(out, data && data.missingNumbers);
  out = enforceAgeFact(out, data && data.age);
  out = enforcePersonalYearFacts(out, data && data.forecastYears);
  return normalizeFinalReport(out);
}

function enforceAgeFact(text, ageValue) {
  const age = Number(ageValue);
  if (!isFinite(age) || age < 0 || age > 120) return String(text || "");
  return String(text || "").replace(/(та\s+(?:одоо|энэ\s+жил)\s+)(\d{1,3})\s*настай/gi, function(match, prefix, n) {
    return Number(n) === age ? match : `${prefix}${age} настай`;
  });
}

function enforcePersonalYearFacts(text, forecastYears) {
  let out = String(text || "");
  const years = forecastYears || [];
  for (let i = 0; i < years.length; i++) {
    const y = years[i];
    const year = Number(y.year);
    const num = Number(y.number);
    if (!year || !num) continue;

    const re = new RegExp(`(${year}\s*он[^\n]{0,120}?Хувийн\s*жил\s*)(\d{1,2})`, "gi");
    out = out.replace(re, function(match, prefix, got) {
      return Number(got) === num ? match : `${prefix}${num}`;
    });
  }
  return out;
}

function generateSectionBody(sectionKey, data, apiKey, systemPrompt, continuity, headers, usageStats, issueHint) {
  const hardFacts = buildHardFactsBlock(data);
  const requiredExcessNumbers = getRequiredExcessNumbers(data);
  const missingLock = buildMissingNumbersLock(data);
  let task = "";
  if (sectionKey === "I") {
    task = `
TASK: Write ONLY the body text for Section I (love DNA, archetype, matrix highlights).
Do NOT write any section header.
Use data:
- Төрсөн огноо: ${data.dob}
- Хүйс: ${data.gender}
- Амьдралын зам: ${data.lifePath} (${data.lifePathName})
- Өдрийн тоо: ${data.dayNumber}
- Матриц: ${data.matrixSummary}
- Илүүдэл тоо: ${JSON.stringify(data.excessNumbers)}
- Матрицад заавал тайлбарлах илүүдэл тоо: ${requiredExcessNumbers.join(", ") || "байхгүй"}
STRICT FOR SECTION I:
- Илүүдэл тоо бүрт тусдаа тайлбар өг.
REFERENCE:
${CONFIG.REFERENCES.PART_1}
`;
  } else if (sectionKey === "II") {
    task = `
TASK: Write ONLY the body text for Section II (single status diagnosis and danger signals).
Do NOT write any section header.
Use data:
- Кармын өр: ${data.karmicDebt}
- Дутуу тоо: ${JSON.stringify(data.missingNumbers)}
- Эрсдэлийн бүлгүүд: ${data.riskProfiles}
FACT LOCK: "Дутуу тоо" зөвхөн ${missingLock} байж болно.
REFERENCE:
${CONFIG.REFERENCES.PART_2}
`;
  } else if (sectionKey === "III") {
    task = `
TASK: Write ONLY the body text for Section III (future partner avatar).
Do NOT write any section header.
Use data:
- Амьдралын зам: ${data.lifePath} (${data.lifePathName})
- Дутуу тоо: ${JSON.stringify(data.missingNumbers)}
- Тохирох өдрийн тоо: ${data.compatibleNumbers}
FACT LOCK: "Дутуу тоо" зөвхөн ${missingLock} байж болно.
REFERENCE:
${CONFIG.REFERENCES.PART_3}
`;
  } else if (sectionKey === "IV") {
    task = `
TASK: Write ONLY the body text for Section IV (3-year love timeline).
Do NOT write any section header.
Use data:
- 3 жилийн прогноз: ${data.forecastText}
RULE: Keep year-number pairs exactly as provided.
REFERENCE:
${CONFIG.REFERENCES.PART_4}
`;
  } else {
    task = `
TASK: Write ONLY the body text for Section V (success strategy and remedies).
Do NOT write any section header.
Use data:
- Амжилтын өнгө: ${data.luckyColor}
- Азын чулуу: ${data.luckyStone}
- Азын тоо: ${data.luckyNumbers}
FACT LOCK: "Дутуу тоо" зөвхөн ${missingLock} байж болно.
REFERENCE:
${CONFIG.REFERENCES.PART_5}
`;
  }

  const prompt = `
${systemPrompt}
${task}
HARD FACTS:
${hardFacts}
STRICT:
- Never contradict HARD FACTS.
- Never write another section header.
- Keep full depth (no short summary mode). Each section body must have multiple substantial paragraphs.
- If QUALITY GAP TO FIX exists, fix those gaps first before expanding style.
QUALITY GAP TO FIX:
${issueHint || "none"}
CONTINUITY:
${continuity || ""}
FORBIDDEN OPENINGS: Never start with "За ойлголоо", "Ойлголоо", "Мэдлээ", "OK".
`;

  const res = callGemini(prompt, apiKey, 3, `REPAIR_SECTION_${sectionKey}`);
  if (usageStats) mergeUsageStats(usageStats, res.usageDetail);
  const cleaned = sanitizeGeneratedPart(res.text, { startHeader: null, stopHeaders: [] });
  const bodyOnly = removeKnownHeaders(cleaned, headers);
  return bodyOnly;
}

function normalizeFinalReport(text) {
  return String(text || "")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function normalizeReportArtifacts(text) {
  let out = normalizeFinalReport(text);
  // Remove invisible characters that sometimes appear as standalone broken symbols.
  out = out.replace(/[\u200B-\u200D\uFEFF]/g, "");
  out = out.replace(/^\s*[\uFE0E\uFE0F]+\s*/gm, "");
  out = out.replace(/^\s*️\s*/gm, "");
  // Remove accidental meta sentence copied from drafting notes.
  out = out.replace(
    /Таны хүссэнээр олон улсын мэргэжлийн нумерологийн жишгээр,\s*зөвхөн тоонд баригдахгүйгээр\s*энергийн мөн чанар,\s*зан төлөв болон таны сэтгэл зүйд үзүүлэх нөлөөллийг эхлээд тайлбарлаж,\s*дараа нь уг энергийг агуулж болзошгүй тоонуудыг болгоомжтойгоор дурдах зарчмаар II Хэсгийг\s*шинэчлэн боловсрууллаа\.?/gi,
    ""
  );
  // Keep list markers deterministic when model emits markdown bullets.
  out = out.replace(/^\s*[\*\-]\s+/gm, "• ");
  out = out.replace(/^\s*•\s+/gm, "• ");
  return normalizeFinalReport(out);
}

// --- HELPER FUNCTIONS ---

function extractGeminiTextFromCandidates(json) {
  const candidates = (json && json.candidates) || [];
  const chunks = [];

  for (let i = 0; i < candidates.length; i++) {
    const content = candidates[i] && candidates[i].content;
    const parts = content && content.parts;
    if (!parts || !parts.length) continue;

    for (let j = 0; j < parts.length; j++) {
      const txt = parts[j] && parts[j].text;
      if (typeof txt === "string" && txt.trim()) chunks.push(txt);
    }
  }

  return chunks.join("\n").trim();
}

function buildGeminiEmptyCandidateReason(json, rawBody) {
  const reasons = [];
  const promptFeedback = (json && json.promptFeedback) || {};
  const candidates = (json && json.candidates) || [];
  const firstCandidate = candidates[0] || {};

  if (promptFeedback.blockReason) reasons.push(`promptBlock=${promptFeedback.blockReason}`);
  if (firstCandidate.finishReason) reasons.push(`finishReason=${firstCandidate.finishReason}`);

  const safety = firstCandidate.safetyRatings || promptFeedback.safetyRatings || [];
  if (Array.isArray(safety) && safety.length) {
    const blockedSafety = [];
    for (let i = 0; i < safety.length; i++) {
      const item = safety[i] || {};
      if (item.blocked === true) {
        blockedSafety.push(`${item.category || "UNKNOWN"}:${item.probability || item.severity || "UNKNOWN"}`);
      }
    }
    if (blockedSafety.length) reasons.push(`safety=${blockedSafety.join(",")}`);
  }

  if (!reasons.length) {
    reasons.push(`raw=${String(rawBody || "").slice(0, 180)}`);
  }

  return reasons.join("; ");
}

function buildShortSafeReference() {
  return [
    "SHORT SAFE REFERENCE:",
    "- Өнгө аяс: тайван, хүндэтгэлтэй, шүүлтгүй, ойлгомжтой монгол хэллэг.",
    "- Prompt-д заасан бүтцийг болон I-V хэсгийн гарчгийг яг мөрдөх.",
    "- Зөвхөн HARD FACTS ба өгөгдсөн тооцоололд тулгуурлах, шинэ баримт зохиохгүй.",
    "- Өдөөн хатгасан, ил задгай, доромжилсон, айдас төрүүлэх хэллэг хэрэглэхгүй.",
    "- Практик, төвийг сахисан зөвлөмжийг тэнцвэртэй найруулгаар өгөх."
  ].join("\n");
}

function buildSafetyFallbackPrompt(originalPrompt) {
  let safe = String(originalPrompt || "");
  const shortSafeRef = buildShortSafeReference();

  // Fallback mode keeps a compact, safer style anchor instead of full long references.
  safe = safe.replace(/REFERENCE STYLE:[\s\S]*$/i, `REFERENCE STYLE:\n${shortSafeRef}`);
  safe = safe.replace(/REFERENCE:[\s\S]*$/i, `REFERENCE:\n${shortSafeRef}`);

  if (!/REFERENCE STYLE:/i.test(safe) && !/REFERENCE:/i.test(safe)) {
    safe += `\n\n${shortSafeRef}\n`;
  }

  safe += `

SAFETY FALLBACK MODE:
- Keep wording respectful, neutral, and non-explicit.
- Avoid provocative/graphic phrasing.
- Preserve all HARD FACTS and requested structure.
`;
  return safe;
}

function callGemini(text, key, attempts = 3, tag = "UNSPECIFIED") {
  if (!key) throw new Error("GENERATE: GEMINI API key байхгүй.");

  const callTag = String(tag || "UNSPECIFIED");
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.GEMINI_MODEL}:generateContent?key=${key}`;
  const baseSafetySettings = [
    { category: "HARM_CATEGORY_HARASSMENT", threshold: "BLOCK_NONE" },
    { category: "HARM_CATEGORY_HATE_SPEECH", threshold: "BLOCK_NONE" },
    { category: "HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold: "BLOCK_NONE" },
    { category: "HARM_CATEGORY_DANGEROUS_CONTENT", threshold: "BLOCK_NONE" }
  ];

  const originalPrompt = String(text || "");
  let activePrompt = originalPrompt;
  let safetyFallbackUsed = false;
  let lastError = "";

  for (let i = 0; i < attempts; i++) {
    try {
      const payload = {
        contents: [{ parts: [{ text: activePrompt }] }],
        generationConfig: {
          temperature: CONFIG.TEMPERATURE,
          maxOutputTokens: 8192 // MAXIMIZED FOR QUALITY
        },
        safetySettings: baseSafetySettings
      };
      const options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      const res = UrlFetchApp.fetch(url, options);
      const status = res.getResponseCode();
      const body = res.getContentText();
      if (status < 200 || status >= 300) {
        lastError = `[${callTag}] HTTP ${status}: ${body.slice(0, 200)}`;
        Utilities.sleep(800);
        continue;
      }

      const json = JSON.parse(body);
      const txt = extractGeminiTextFromCandidates(json);
      if (txt) {
        const usageMeta = json.usageMetadata || {};
        const promptTokens = Number(usageMeta.promptTokenCount || usageMeta.inputTokenCount || 0);
        const outputTokens = Number(usageMeta.candidatesTokenCount || usageMeta.outputTokenCount || 0);
        const totalTokens = Number(usageMeta.totalTokenCount || (promptTokens + outputTokens));
        
        const clean = txt
          .replace(/\-\-\-/g, "")
          .replace(/###/g, "")
          .replace(/\(Үргэлжлэл\)/g, "")
          .replace(/ТА-ны/g, "Таны")
          .replace(/Тануулга/g, "Тайлал")
          .trim();
          
        return {
          text: clean,
          usage: totalTokens,
          usageDetail: {
            totalTokens: totalTokens,
            inputTokens: promptTokens,
            outputTokens: outputTokens,
            requests: 1
          }
        };
      }

      const promptBlock = (json && json.promptFeedback && json.promptFeedback.blockReason) || "";
      const finishReason = (json && json.candidates && json.candidates[0] && json.candidates[0].finishReason) || "";
      if (!safetyFallbackUsed && (promptBlock === "PROHIBITED_CONTENT" || finishReason === "SAFETY")) {
        const fallbackPrompt = buildSafetyFallbackPrompt(originalPrompt);
        if (fallbackPrompt && fallbackPrompt !== activePrompt) {
          activePrompt = fallbackPrompt;
          safetyFallbackUsed = true;
          lastError = `[${callTag}] response blocked (${promptBlock || finishReason}). safety fallback prompt руу шилжлээ.`;
          console.warn(`[${callTag}] Gemini blocked, using safety fallback prompt.`);
          Utilities.sleep(700);
          continue;
        }
      }

      lastError = `[${callTag}] response candidates хоосон. ${buildGeminiEmptyCandidateReason(json, body)}`;
      console.warn(`[${callTag}] Gemini empty candidate:`, lastError);
    } catch (e) {
      lastError = String(e.message || e);
      console.error(`[${callTag}] Gemini attempt ${i + 1} failed:`, e);
    }
    Utilities.sleep(1000);
  }

  throw new Error(`GENERATE: Gemini амжилтгүй. ${lastError}`);
}

// --- MATH LOGIC ---
function extractJsonBlock(text) {
  const cleaned = String(text || "").replace(/```json/gi, "").replace(/```/g, "").trim();
  const first = cleaned.indexOf("{");
  const last = cleaned.lastIndexOf("}");
  if (first === -1 || last === -1 || first >= last) return {};
  return JSON.parse(cleaned.slice(first, last + 1));
}

function normalizeDateValue(input) {
  const s = String(input || "").trim();
  // Year-month-day with separators, with numeric token boundaries.
  let m = s.match(/(?:^|[^\d])(\d{4})(?:[\.\-\/\s]+)(\d{1,2})(?:[\.\-\/\s]+)(\d{1,2})(?!\d)/);
  if (m) {
    const y = normalizeYearToken(m[1]);
    const mo = Number(m[2]);
    const d = Number(m[3]);
    if (y && isValidDate(y, mo, d)) {
      return `${y}.${String(mo).padStart(2, "0")}.${String(d).padStart(2, "0")}`;
    }
  }

  // Two-digit year format.
  m = s.match(/(?:^|[^\d])(\d{2})(?:[\.\-\/\s]+)(\d{1,2})(?:[\.\-\/\s]+)(\d{1,2})(?!\d)/);
  if (m) {
    const yy = Number(m[1]);
    const y = yy > 25 ? 1900 + yy : 2000 + yy;
    const mo = Number(m[2]);
    const d = Number(m[3]);
    if (isValidDate(y, mo, d)) {
      return `${y}.${String(mo).padStart(2, "0")}.${String(d).padStart(2, "0")}`;
    }
  }

  // Fuzzy parser for human-typed text:
  // "1996 10 sariin 30nd torson", "11996 10 30", "19961030"
  const numericTokens = s.match(/\d+/g) || [];
  if (numericTokens.length) {
    // 8-digit compact date: YYYYMMDD
    for (let i = 0; i < numericTokens.length; i++) {
      const tk = numericTokens[i];
      if (tk.length === 8) {
        const y = Number(tk.slice(0, 4));
        const mo = Number(tk.slice(4, 6));
        const d = Number(tk.slice(6, 8));
        if (isValidDate(y, mo, d)) {
          return `${y}.${String(mo).padStart(2, "0")}.${String(d).padStart(2, "0")}`;
        }
      }
    }

    // Prefer year-first sequences.
    for (let i = 0; i <= numericTokens.length - 3; i++) {
      const y = normalizeYearToken(numericTokens[i]);
      if (!y) continue;
      const mo = Number(numericTokens[i + 1]);
      const d = Number(numericTokens[i + 2]);
      if (isValidDate(y, mo, d)) {
        return `${y}.${String(mo).padStart(2, "0")}.${String(d).padStart(2, "0")}`;
      }
    }

    // Secondary fallback: day-month-year sequence.
    for (let i = 2; i < numericTokens.length; i++) {
      const y = normalizeYearToken(numericTokens[i]);
      if (!y) continue;
      const d = Number(numericTokens[i - 2]);
      const mo = Number(numericTokens[i - 1]);
      if (isValidDate(y, mo, d)) {
        return `${y}.${String(mo).padStart(2, "0")}.${String(d).padStart(2, "0")}`;
      }
    }
  }
  return null;
}

function normalizeYearToken(token) {
  const raw = String(token || "").replace(/\D/g, "");
  if (!raw) return null;

  if (raw.length === 4) {
    const y = Number(raw);
    if (y >= 1900 && y <= 2099) return y;
  }

  // Common typo: accidental leading digit, e.g. 11996 -> 1996
  if (raw.length === 5) {
    const head = Number(raw.slice(0, 4));
    if (head >= 1900 && head <= 2099) return head;
    const tail = Number(raw.slice(1));
    if (tail >= 1900 && tail <= 2099) return tail;
  }

  // Long noisy token: scan every 4-digit window and pick a plausible year.
  if (raw.length > 5) {
    for (let i = 0; i <= raw.length - 4; i++) {
      const y = Number(raw.slice(i, i + 4));
      if (y >= 1900 && y <= 2099) return y;
    }
  }

  if (raw.length === 2) {
    const yy = Number(raw);
    return yy > 25 ? 1900 + yy : 2000 + yy;
  }

  return null;
}

function isValidDate(y, m, d) {
  if (m < 1 || m > 12 || d < 1 || d > 31) return false;
  const dt = new Date(y, m - 1, d);
  return dt.getFullYear() === y && dt.getMonth() + 1 === m && dt.getDate() === d;
}

function normalizeGender(raw, aiGender) {
  const direct = detectGenderFromRaw(raw);
  if (direct) return direct;
  const g = String(aiGender || "").trim();
  if (g === "Эрэгтэй" || g === "Эмэгтэй") return g;
  // Strict mode: wrong gender inference is worse than retry/review.
  throw new Error("NORMALIZE: Хүйс тодорхойгүй байна. (ЭРЭГТЭЙ/ЭМЭГТЭЙ мэдээлэл шаардлагатай)");
}

function detectGenderFromRaw(raw) {
  const text = String(raw || "").toLowerCase();
  if (text.includes("эрэгтэй") || text.includes("eregtei") || /(^|\s)эр(\s|$)/.test(text)) return "Эрэгтэй";
  if (text.includes("эмэгтэй") || text.includes("emegtei") || /(^|\s)эм(\s|$)/.test(text)) return "Эмэгтэй";
  return null;
}

function normalizeName(raw, aiName) {
  const fromAi = String(aiName || "").trim();
  if (fromAi && !/\d/.test(fromAi) && fromAi.length <= 40) return fromAi;

  const cleaned = String(raw || "").replace(/[\d\.\-\/]/g, " ").trim();
  const parts = cleaned.split(/\s+/).filter(Boolean);
  for (let i = 0; i < parts.length; i++) {
    const p = parts[i];
    if (!["эр", "эм", "эрэгтэй", "эмэгтэй"].includes(p.toLowerCase())) {
      return p;
    }
  }
  return "Эрхэм";
}

function getGenderStyle(gender) {
  const key = String(gender || "").trim();
  return NUMEROLOGY.GENDER_STYLE_MAP[key] || NUMEROLOGY.GENDER_STYLE_MAP["Эмэгтэй"];
}

function escapeRegExp(input) {
  return String(input).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function applyGenderLanguageGuard(text, gender) {
  let out = String(text || "");
  const style = getGenderStyle(gender);
  const rules = style.replaceRules || [];

  for (let i = 0; i < rules.length; i++) {
    const rule = rules[i];
    const phrase = escapeRegExp(rule.from);
    // Replace only standalone phrases to avoid "тэр хүн" -> "тэмэгтэй хүн" corruption.
    const re = new RegExp(`(^|[^\\p{L}\\p{N}])(${phrase})(?=$|[^\\p{L}\\p{N}])`, "giu");
    out = out.replace(re, `$1${rule.to}`);
  }

  // Always keep honorifics neutral to avoid tone mismatch.
  out = out.replace(/ноён/gi, "хүн");
  out = out.replace(/хатагтай/gi, "хүн");
  out = out.replace(/үйлчлүүлэгч/gi, "хүн");
  return out;
}

function calculateAge(y, m, d) {
  const now = new Date();
  let age = now.getFullYear() - y;
  const currentMonth = now.getMonth() + 1;
  const currentDay = now.getDate();
  if (currentMonth < m || (currentMonth === m && currentDay < d)) age--;
  return age;
}

function createPdf(name, content, templateId) {
  const folderId = PropertiesService.getScriptProperties().getProperty("FOLDER_ID");
  if (!folderId) throw new Error("PDF: FOLDER_ID байхгүй.");
  if (!templateId) throw new Error("PDF: TEMPLATE_ID байхгүй.");

  const copy = DriveApp.getFileById(templateId).makeCopy(`${name} - ${CONFIG.PRODUCT_NAME}`);
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();
  
  // 1. FILL CONTENT
  // Case-insensitive placeholders: {{NAME}}/{{name}}, {{REPORT}}/{{report}}
  body.replaceText("(?i)\\{\\{\\s*name\\s*\\}\\}", name);
  body.replaceText("(?i)\\{\\{\\s*report\\s*\\}\\}", content);
  
  // 2. FORMATTING ENGINE (CLASSIC BOLD + CLEAN LISTS)
  processFormatting(body);

  // 3. PARAGRAPH STYLING
  const paragraphs = body.getParagraphs();
  for (let i = 0; i < paragraphs.length; i++) {
    paragraphs[i].setLineSpacing(1.5);
    paragraphs[i].setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
  }
  
  doc.saveAndClose();
  
  const pdf = copy.getAs(MimeType.PDF);
  const pdfFile = DriveApp.getFolderById(folderId).createFile(pdf); 
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  copy.setTrashed(true);
  return {
    url: pdfFile.getUrl(),
    fileId: pdfFile.getId()
  };
}

// --- FORMATTING ENGINE (CLASSIC BOLD + BULLETS) ---
function processFormatting(body) {
  // 1. Handle BOLD: **text** -> Bold text without markers
  var foundElement = body.findText("\\*\\*(.*?)\\*\\*");
  while (foundElement != null) {
    var foundText = foundElement.getElement().asText();
    var start = foundElement.getStartOffset();
    var end = foundElement.getEndOffsetInclusive();

    if (end - start >= 3) {
      foundText.setBold(start + 2, end - 2, true);
      foundText.deleteText(end - 1, end);
      foundText.deleteText(start, start + 1);
    } else {
      break;
    }

    foundElement = body.findText("\\*\\*(.*?)\\*\\*");
  }
  
  // 2. Handle BULLETS: "* text", "- text", "• text" -> Bullet point
  var paras = body.getParagraphs();
  for (var i = 0; i < paras.length; i++) {
    var p = paras[i];
    var text = p.getText();
    var trimmed = text.trim();
    if (/^(?:\*|-|•)\s+/.test(trimmed) && !trimmed.startsWith("**")) {
      p.setText(text.replace(/^\s*(?:\*|-|•)\s+/, "")); // Remove marker
      p.setGlyphType(DocumentApp.GlyphType.BULLET); // Make it a bullet
    }
  }
}

function sendManyChat(subscriberId, pdfUrl, name, token) {
  if (!token) throw new Error("DELIVERY: ManyChat token байхгүй.");
  if (!subscriberId) throw new Error("DELIVERY: subscriber_id хоосон.");
  if (!pdfUrl) throw new Error("DELIVERY: PDF URL хоосон.");

  const text = renderDeliveryMessage(CONFIG.DELIVERY_MESSAGE, name || "Эрхэм", pdfUrl);

  const payload = {
    subscriber_id: String(subscriberId).trim(),
    data: {
      version: "v2",
      content: {
        messages: [
          {
            type: "text",
            text: text
          }
        ]
      }
    }
  };

  const res = UrlFetchApp.fetch("https://api.manychat.com/fb/sending/sendContent", {
    method: "post",
    headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const status = res.getResponseCode();
  const body = res.getContentText();
  if (status < 200 || status >= 300) {
    throw new Error(`DELIVERY HTTP ${status}: ${body.substring(0, 200)}`);
  }
}

function renderDeliveryMessage(template, name, pdfUrl) {
  const safeName = String(name || "Эрхэм").trim();
  const safeUrl = String(pdfUrl || "").trim();
  let out = String(template || "");

  out = out
    .replace(/\{\{?\s*NAME\s*\}?\}/gi, safeName)
    .replace(/\{\{?\s*URL\s*\}?\}/gi, safeUrl)
    .replace(/%NAME%/gi, safeName)
    .replace(/%URL%/gi, safeUrl);

  // Last-resort cleanup for malformed braces.
  out = out
    .replace(/\{+\s*NAME\s*\}+/gi, safeName)
    .replace(/\{+\s*URL\s*\}+/gi, safeUrl);

  return out;
}
