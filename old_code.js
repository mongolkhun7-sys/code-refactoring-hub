/**
 * PRODUCT: MONGOLIAN COUPLE ASTROLOGY (HARMONY V18.9)
 * CONCEPT: "GOLD STANDARD" - Header Section Added + 6-Step Split
 * AUTHOR: Jules (Strict Adherence to User Requirements)
 * MODEL: gemini-2.5-flash
 ****************************************************************************************/

const CONFIG = {
  // --- SYSTEM CONFIGURATION ---
  VERSION: "v18.9-HarmonyGold-HeaderAdded",
  PRODUCT_NAME: "Хосын Нийцлийн Алтан Тайлан",
  SHEET_NAME: "Sheet1",
  BATCH_SIZE: 2, 
  GEMINI_MODEL: "gemini-2.5-flash", 
  TEMPERATURE: 0.6, 

  // --- COLUMN MAPPING (0-based Index) ---
  COLUMNS: {
    NAME: 0,      // A: User Name
    ID: 1,        // B: Subscriber ID
    INPUT: 2,     // C: Raw Input (Gender DOB1 DOB2)
    PDF: 3,       // D: Output PDF URL
    STATUS: 4,    // E: Processing Status
    TOKEN: 5,     // F: Token Usage
    DEBUG: 6,     // G: Debug Data
    DATE: 7,      // H: Timestamp
    VER: 8,       // I: Version
    ERROR: 9      // J: Error
  },

  MAX_EXECUTION_TIME: 360000, 
  SAFETY_BUFFER: 60000,       

  // ==================================================================================
  // 📜 FULL "GOLD STANDARD" REFERENCE TEXTS (STYLE GUIDE ONLY)
  // ==================================================================================
  REFERENCES: {
    PART_1: `1-р хэсэг. ЯЗГУУР МӨН ЧАНАР: Жилүүдийн Ивээл ба Харш
🕯️ Эхлээд та хоёрын төрсөн жилийн мөн чанарыг тодорхойлбол эрэгтэй нь 1996 оны "Гал Хулгана" жилтэй бол эмэгтэй нь 1999 оны "Шороо Туулай" жилтэй байна. Монгол зурхайн ухаанд энэхүү хоёр жил нь хоорондоо харшлах "Түйтгэр" болон "Их харш" харилцаанд ордоггүй тул суурь харилцаа нь тайван, бие биедээ дарамт учруулахгүй байх өгөгдөлтэй. Гэхдээ зөвхөн амьтан бус, махбодын харилцааг гүнзгийрүүлэн харвал дараах зүй тогтол ажиглагдаж байна.
🐁 Эрэгтэй нь Хулгана жилтэй тул байгалийн өгөгдөл нь аливааг хурааж хуримтлуулах, жижиг зүйлийг ч анзаарч харах, сэргэлэн сийрэг ухааныг илтгэдэг. Хулгана нь шөнийн амьдралтай, соргог амьтан учраас эрэгтэй нь гэр бүлийнхээ төлөө байнга санаа зовниж, ирээдүйгээ тооцоолж, гэнэтийн эрсдэлээс хамгаалах хариуцлагыг өөр дээрээ үүрэх хандлагатай байдаг. Учир нь хулгана жил нь өөрөө "эхлэл"-ийн бэлгэдэл тул аливаа шинэ зүйлийг санаачлагч, гэр бүлийнхээ санхүүгийн аюулгүй байдлын манаач нь байх болно. Түүний "Гал" махбод нь түүнийг дотооддоо дүрэлзсэн хүсэл тэмүүлэлтэй, зорилгодоо хүрэхийн төлөө тууштай тэмцэгч гэдгийг харуулж байна.
🐇 Эмэгтэй нь Туулай жилтэй тул төрөлхийн зөөлөн дөлгөөн, амар амгаланг эрхэмлэдэг, зөрчилдөөнөөс зугтдаг мөн чанартай. Туулай нь үсрэх чадвартай, хурдан шалмаг боловч маш болгоомжтой амьтан учраас эмэгтэй нь аливаа шийдвэр гаргахдаа яардаггүй, бүх талыг нь нягталж байж хөдөлдөг нямбай чанарыг харилцаанд авчирдаг. Түүний "Шороо" махбод нь тогтвортой байдал, тэсвэр тэвчээр, үржил шимийг илтгэх тул тэрээр гэр бүлдээ түшиг тулгуур болж, эрсдэлтэй алхмуудыг зөөлрүүлж, бодит хөрсөн дээр буулгах үүргийг гүйцэтгэдэг байна.
❤️ Та хоёрын жилийн махбодын харилцааг шинжвэл маш онцгой сайн зохицол харагдаж байна. Учир нь эрэгтэйн "Гал" махбод нь эмэгтэйн "Шороо" махбодыг төрүүлдэг (Галнаас үнс буюу шороо үүсдэг) зүй тогтолтой. Үүнийг зурхайд "Эх хүүгийн барилдлага" гэж үздэг бөөд энэ нь эрэгтэй нь эмэгтэйгээ өөрийн эрч хүчээр тэтгэж, хайрлаж, урам зориг өгч байдаг гэсэн үг юм. Эрэгтэй нь санаачилж, гал мэт дүрэлзэхэд эмэгтэй нь тэр энергийг хүлээн авч, шороо мэт бат бэх суурийг бүрдүүлнэ гэсэн үг. Иймээс та хоёрын харилцаанд эрэгтэй нь илүү халамжтай, өгөгч байр суурьтай байх бол эмэгтэй нь хүлээн авагч, тогтворжуулагч байх нь байгалийн хуулиар заяагдсан зөв зохицол юм.
🗝️ Дүгнэж хэлэхэд Хулгана, Туулай хоёр нь зан чанарын хувьд өөр (нэг нь идэвхтэй цуглуулагч, нөгөө нь болгоомжтой ажиглагч) боловч яг энэ ялгаатай байдал нь нэгнийхээ дутууг нөхөх боломжийг олгоно. Эрэгтэйн овсгоо самбаа, эмэгтэйн дипломат зөөлөн зан хоёр нийлэхэд гэр бүлийн дотоод уур амьсгал маш тэнцвэртэй байх болно. Гол анхаарах зүйл нь Хулгана жилтэн хэт жижиг зүйлд анхаарч шүүмжлэх гээд байдаг талтай бол Туулай жилтэн гомдомхой байдаг тул эрэгтэй нь шүүмжлэлээ багасгаж, эмэгтэй нь аливааг хэт хүндээр тусгаж авахгүй байх нь та хоёрын язгуур мөн чанарын зохицлыг төгс болгох түлхүүр юм.`,

    PART_2: `2-р хэсэг. ЭРЧИМ ХҮЧНИЙ НИЙЛЭМЖ: Махбодын Зохицол
🔥 Энэ хэсэгт бид та хоёрын зөвхөн төрсөн жил бус, тухайн жилийг удирдаж буй байгалийн суурь энерги буюу "Махбод"-ын харилцааг гүнзгийрүүлэн авч үзнэ. Эрэгтэй нь "Гал" махбодтой, эмэгтэй нь "Шороо" махбодтой байгаа нь энергийн солилцооны хувьд "Тэтгэгч ба Тэтгүүлэгч" гэсэн маш таатай, зөв урсгалыг үүсгэдэг. Яагаад гэвэл байгаль дээр гал дүрэлзэж дууссаны дараа үнс буюу шороо болж хувирдаг зүй тогтолтой тул эрэгтэйн энерги нь эмэгтэйгээ байнга тэжээж, өөд нь татаж, амьдралын эрч хүчээр сэлбэж байдаг гэсэн үг юм. Ийм учраас танай гэр бүлд эрэгтэй нь гол хөдөлгөгч хүч, санаачлагч байж, харин эмэгтэй нь тэрхүү энергийг бодит үр дүн болгон хувиргадаг "сав" нь болдог.
⚡ Сэтгэл хөдлөлийн хэмнэлийн хувьд та хоёр бие биеэ нөхөх зарчмаар ажиллана. "Гал" махбодтой эрэгтэй нь сэтгэл хөдлөл өндөртэй, түргэн ууртай боловч маш хурдан тайвширдаг, цайлган зантай байх магадлалтай. Харин "Шороо" махбодтой эмэгтэй нь аливааг дотроо боловсруулдаг, удаан боддог, сэтгэл хөдлөлөө тэр бүр ил гаргаад байдаггүй тогтуун шинжтэй. Тиймээс эрэгтэй нь дүрэлзээд уурлахад эмэгтэй нь түүнийг шороо мэт дарж, тайвшруулж чаддаг. Гэхдээ анхаарах зүйл нь Гал махбод хэт дүрэлзвэл шороог хурайшуулж хатууруулдаг шиг, эрэгтэй нь хэт их давамгайлж, сэтгэл хөдлөлөөр дайрвал эмэгтэй нь дотроо гомдол тээж, зөрүүдлэн дуугаа хураах эрсдэлтэйг санах хэрэгтэй.
💊 Эрүүл мэндийн болон бие махбодын нийцлийн тухайд энэ хоёр махбод нь бие биедээ "эм" болдог. Гал нь зүрх, цусны эргэлтийг хариуцдаг бол Шороо нь ходоод, дэлүү, булчинг хариуцдаг. Эрэгтэйн идэвхтэй, халуун энерги нь эмэгтэйн удаан, хүйтэн байж болох энергийг бүлээцүүлж, цусны эргэлтийг нь сайжруулна. Харин эмэгтэйн тогтвортой энерги нь эрэгтэйн хэт их ачаалласан мэдрэлийн системийг амраах үйлчилгээтэй. Иймээс та хоёр хамт байхдаа бие биеийнхээ ядаргааг тайлж, нэг нь нөгөөдөө эрч хүч өгдөг биологийн зөв хэмнэлтэй хосууд юм.
🌱 Эцэст нь дүгнэхэд, энэхүү махбодын зохицол нь та хоёрын харилцааг "Эх хүүгийн барилдлага" шиг нандин болгож байна. Энд "Эх" нь Гал (Эрэгтэй), "Хүү" нь Шороо (Эмэгтэй) болж байгаа юм (Галнаас Шороо төрдөг тул). Эрэгтэй нь хайр халамжаа урсгаж, эмэгтэй нь түүнийг шингээж авснаар сэтгэл ханамжтай, аюулгүй мэдрэмжийг авдаг. Ганцхан анхаарах зүйл бол эрэгтэй нь өөрийнхөө галыг байнга зарцуулаад байвал амархан ядарч магадгүй тул эмэгтэй нь хариу талархал, халамжаар түүний галыг унтраалгүй дэмжиж байх нь энэхүү төгс зохицлыг хадгалах гол үндэс болно.`,

    PART_3: `3-р хэсэг. ОЮУН САНААНЫ ХОЛБООС: Мэнгэний Таарамж
🧠 Энэ хэсэгт та хоёрын сэтгэлгээний онцлог, амьдралыг үзэх үзэл болон оюун санааны нийцлийг "Мэнгэ"-ний зурхайгаар тайлбарлая. Тооцоолж үзвэл эрэгтэй нь "4 ногоон" мэнгэтэй бол эмэгтэй нь "5 шар" мэнгэтэй байна. Энэ хоёр мэнгийн нийлэмж нь Монгол зурхайд "Эмч, домч" буюу бие биеэ анагаах барилдлага гэж буудаг. Яагаад ийм нэртэй бууж байгааг болон энэ нь бодит амьдрал дээр юу гэсэн үг болохыг дор задалж тайлбарлая.
🌲 Эрэгтэйн "4 ногоон" мэнгэ нь байгалийн мөн чанараараа "Мод" махбодыг илтгэдэг. Мод нь ургаж дээшилдэг, мөчирлөн тэлдэг шинжтэй тул эрэгтэй нь оюун санааны хувьд маш чөлөөт сэтгэлгээтэй, шинийг эрэлхийлэгч, нэг газартаа тогтох дургүй, үргэлж өөрийгөө хөгжүүлэх хүсэлтэй байдаг. Тэрээр аливаа асуудалд уян хатан хандаж, салхи мэт хувирч чаддаг сэргэлэн зантай.
⛰️ Эмэгтэйн "5 шар" мэнгэ нь "Шороо" махбодыг илтгэдэг бөгөөд энэ нь бүх тооны төв дунд оршдог "Хаан мэнгэ" гэгддэг. Иймээс эмэгтэй нь оюун санааны хувьд маш тогтвортой, өөрийн гэсэн бат бөх үзэл бодолтой, аливаад голч байр сууринаас ханддаг, зарчимч хүн байх магадлалтай. Тэрээр хийсвэр зүйлээс илүү бодитой, гарт баригдах үр дүнг чухалчилдаг.
⚕️ Та хоёрын харилцааг "Эмч домч" гэж нэрлэсний учир нь: Мод нь шороон дээр үндэслэж ургадаг зүй тогтолтой. Үүнтэй адил эрэгтэйн тэлэх, өсөх хүсэл тэмүүлэл нь эмэгтэйн бат бөх, тогтвортой суурин дээр тогтож байж амжилт олдог гэсэн үг юм. Энэ харилцаа нь яг л эмчилгээ шиг үйлчилдэг. Хэн нэгэн нь сэтгэлээр унах, алдаа гаргах, эсвэл өвдөх үед нөгөөх нь түүнийг засаж, тордож, эдгээдэг онцгой үйлийн үртэй хосууд юм. Та хоёр бие биенийхээ сэтгэлийн шархыг хамгийн сайн ойлгож, эмчилж чадна.
⚖️ Гэхдээ анхаарах нэг нарийн зүйл бий. Байгаль дээр мод нь шорооны шим тэжээлийг сорж ургадаг тул эрэгтэй нь өөрийн үзэл бодлоо эмэгтэйдээ тулгах, түүнийг өөрийнхөөрөө байлгах гэж оролдох хандлага гарч мэднэ. Эмэгтэй нь "5 шар" буюу хүчтэй мэнгэтэй тул үүнийг нь эсэргүүцэж магадгүй. Тиймээс эрэгтэй нь эмэгтэйнхээ тогтвортой чанарыг хүндэлж, харин эмэгтэй нь эрэгтэйнхээ шинэлэг санааг дэмжиж байвал энэ нь хамгийн урт удаан, бие биедээ ээлтэй, "анагаагч" харилцаа болж төлөвшинө.`,

    PART_4: `4-р хэсэг. АМЬДРАЛЫН ХЭМНЭЛ: Суудлын Өнцөг & Гал Голомт
🏠 Энэ хэсэг нь та хоёрын өдөр тутмын ахуй амьдрал, гэр орны уур амьсгал, санхүүгийн удирдлага болон гал тогоонд хэрхэн зохицохыг харуулна. Энэ бол хайр дурлалаас илүүтэй хамтран амьдрах чадварыг хэмждэг хэсэг юм.
🌪️ Эрэгтэй нь Цон суудалтай. Энэ нь зүүн урд зүгийг эзэлдэг, хавар, ургаж буй мод болон салхины бэлгэдэлтэй. Ийм суудалтай хүн гэр бүлдээ шинэ урсгал, өөрчлөлт, хөдөлгөөнийг авчирдаг. Тэрээр нэг газраа удаан суух дургүй, гэрийн тавилгаа байршлыг нь солих, эсвэл мөнгөө эргэлдүүлэх, аялах сонирхолтой хувьсамтгай хэв маягийг илтгэнэ.
⛰️ Эмэгтэй нь Гэн суудалтай. Энэ нь зүүн хойд зүгийг эзэлдэг, уул болон шороон махбодыг илтгэнэ. Уул хөдөлдөггүй шиг эмэгтэй нь гэр бүлдээ тогтвортой байдал, хуучин уламжлал, хадгаламж, бат бөх байдлыг эрхэмлэдэг. Тэрээр гэртээ байх дуртай, эрсдэлээс айдаг, юмыг яг байран нь байлгахыг хүсдэг суурин хэв маягтай.
⚠️ Суудлын нийцлийг шинжвэл Хорт буюу Анхаарах харилцаа үүсэж байна. Энд бид бодит үнэнийг хэлэх ёстой. Өмнөх жил болон махбодын зохицол маш сайн байсан ч, суудлын хувьд та хоёр Хорт гэх харилцаанд бууж байна. Яагаад ийм нэртэй вэ гэвэл Байгалийн хуулиар Мод буюу Эрэгтэй нь Шороог буюу Эмэгтэйг үндсээрээ хага ярж ургадаг. Салхи буюу Эрэгтэй нь Уулыг буюу Эмэгтэйг тойрч үлээгээд хүчгүйддэг эсвэл элээдэг.
🚧 Энэ нь амьдрал дээр юу гэсэн үг вэ гэхээр Эрэгтэй нь ямар нэг шинэ зүйл хийх гэхээр эмэгтэй нь хэрэггүй, эрсдэлтэй гэж хориглох, эсвэл эрэгтэй нь эмэгтэйгээ хэт удаан, хөшүүн байна гэж шүүмжлэх хандлага гарч болзошгүй. Санхүүгийн тал дээр эрэгтэй нь үрэх эсвэл эргэлдүүлэх, эмэгтэй нь хадгалах эсвэл дарах гэсэн хоёр өөр туйл руу зүтгэж магадгүй.
🔥 Гэхдээ бүү санаа зов, учир нь Зурхай бол онош, харин Дом бол эмчилгээ юм. Та хоёрын энэхүү суудлын зөрчлийг арилгах гол гүүр бол ГАЛ махбод юм. Мод буюу Эрэгтэй шатаж Гал болно, Гал нь унтраад Шороо буюу Эмэгтэй болно. Ингэж дундаа Гал махбодыг оруулж ирснээр зөрчил биш, эргэлт үүснэ.
🕯️ Иймээс та хоёр гал тогоо болон унтлагын өрөөндөө улаан өнгийн декор, гэрэлтүүлэг, эсвэл лаа байнга байлгах нь сайн. Мөн хурим найр хийх, гэр бүл болохдоо Галсан буюу Гал тахих зан үйлийг заавал хийлгэх хэрэгтэй. Энэ нь та хоёрын ахуйн зөрчлийг шатааж, эв эеийг бэхжүүлнэ.
🗝️ Дүгнэж хэлэхэд хэдийгээр ахуйн амьдрал дээр үзэл бодлын зөрүү гарах магадлалтай ч, та хоёрын сэтгэл зүрхний холбоо маш сайн тул энэ нь зөвхөн жижиг маргаан төдий байх болно. Гол нь эрэгтэй нь эмэгтэйнхээ хадгаламжид гар хүрэхгүй байх, эмэгтэй нь эрэгтэйнхээ шинэ санааг шууд няцаахгүй байх дүрмийг гэртээ тогтоовол энэ суудлын харшийг давуу тал болгож чадна.`,

    PART_5: `5-р хэсэг. ҮЙЛИЙН ҮР БА ИРЭЭДҮЙН ЗУРАГЛАЛ Нэгдсэн Дүгнэлт
🔮 Энэ хэсэгт бид өмнөх дөрвөн хэсгийн мэдээллийг базаж, та хоёрын ирээдүйн ерөнхий зураглал болон үйлийн үрийн холбоог эцэслэн дүгнэж байна. Та хоёрын харилцааг нэг үгээр тодорхойлбол Газар тэнгэрийн савлагаатай ч, ган бат гагнагдсан сүнснүүд гэж хэлж болно. Учир нь та хоёрын суурь өгөгдөл буюу Жил, Мэнгэний зохицол нь маш сайн буюу сүнс, сэтгэл таарч байгаа бол Суудлын зохицол нь таарамж муутай буюу ахуй амьдрал дээр зөрчилтэй байгаа нь ийм дүгнэлтэд хүргэж байна.
🔗 Та хоёр яагаад бие биеэ олсон бэ гэвэл үйлийн үрийн үүднээс харвал энэ нь санамсаргүй учрал биш юм. Эмч домч мэнгэний барилдлага болон Эх хүү жилийн барилдлага нь та хоёрыг урьд насандаа бие биедээ өртэй, эсвэл дуусгаагүй үйл хэрэгтэй байсныг илтгэдэг. Энэ амьдралдаа та хоёр бие биеэ эмчилж, дутуугаа нөхөж, нэг нь нөгөөгөө өсгөн хөгжүүлэх гэж учирсан байна. Эрэгтэй нь эмэгтэйдээ урам зориг, амьдралын оч бэлэглэх үүрэгтэй бол, эмэгтэй нь эрэгтэйдээ амрах орон зай, тогтвортой байдлыг бэлэглэх гэр нь болох учиртай.
🛠️ Ирээдүйд юуг анхаарах вэ гэвэл харилцааны хамгийн том сорилт нь сэтгэл хөрөхдөө биш, харин өдөр тутмын жижиг зөрчил маргаануудад байх болно. Өмнө дурдсан Хорт суудлын нөлөөгөөр ахуйн асуудал, мөнгө санхүү, гэр бүлийн дэг жаяг дээр санал зөрөлдөх магадлал өндөр. Гэхдээ үүнийг бид таарахгүй юм байна гэж хүлээж авах биш, харин энэ бол бидний шийдэх ёстой карма гэж харах хэрэгтэй. Байгалийн хуулиар Мод буюу Эрэгтэй үндэслэхийн тулд Газар буюу Эмэгтэй хэрэгтэй, Газар үржил шимтэй байхын тулд Мод хэрэгтэй байдаг шиг та хоёр бие биегүйгээрээ дутуу мэдрэмж авах тул салалт ховор тохиодог бат бөх хосууд юм.
🔥 Аз жаргалын түлхүүр буюу Дом нь та хоёрын харилцааг аврах, тэтгэх гол элемент бол ГАЛ юм. Яагаад гэвэл эрэгтэйн жилийн махбод Гал, мөн Мод суудал нь галыг тэтгэдэг, тэр гал нь эргээд эмэгтэйн Шороо махбодыг үүсгэдэг зүй тогтолтой. Иймд харилцаагаа үргэлж халуун дулаан байлгах, бие биедээ байнга сэтгэлээ илэрхийлэх, романтик байдлыг гээхгүй байх нь та хоёрын хувьд зүгээр нэг зугаа биш, харин гэр бүлээ хамгаалах хамгаалалт юм. Хэрүүл маргаан гарвал удаан дуугаа хураалгүй буюу шороо чанараа гаргалгүй, шууд ярилцаж шатааж, шийдвэрлэж буюу гал чанараа гаргаж байхыг зөвлөж байна.
💎 Эцсийн үнэлгээгээр та хоёрын харилцаа 85 хувьтай буюу Маш сайн гэж дүгнэгдэж байна. Хэдийгээр ахуйн зөрчил байж болох ч, оюун санаа, бие махбод, сэтгэл зүрхний нийцэл маш өндөр тул та хоёр хамтдаа урт удаан, жаргалтай амьдрах бүрэн боломжтой. Эрэгтэйн овсгоо самбаа, эмэгтэйн хатан ухаан хоёр нийлсэн байхад давахгүй даваа гэж үгүй ээ. Сайн сайхныг ерөөе!`
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
  
  // Strict Mapping: (Year - 1900) % 12
  ANIMAL_MAP: [
    "Хулгана", "Үхэр", "Бар", "Туулай", "Луу", "Могой", 
    "Морь", "Хонь", "Бич", "Тахиа", "Нохой", "Гахай"
  ],

  // Element by Last Digit
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

  SUUDAL_GROUPS: {
    "Хам": "East", "Зэн": "East", "Цон": "East", "Ли": "East",
    "Хон": "West", "Төмөр": "West", "Дүй": "West", "Гэн": "West"
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
  },

  DELIVERY_MESSAGE: `Сайн байна уу? 🔮\n\nТа хоёрын "Хосын Нийцлийн Алтан Тайлан" бэлэн боллоо.\n\nФайл: {{URL}}\n\n(Татаж аваад хадгалаарай)`,
  
  INTRO_TEXTS: {
    PART_1: "Эхлээд та хоёрын төрсөн жилийн мөн чанарыг тодорхойлбол...",
    PART_2: "Энэ хэсэгт бид та хоёрын зөвхөн төрсөн жил бус, тухайн жилийг удирдаж буй байгалийн суурь энерги буюу 'Махбод'-ын харилцааг гүнзгийрүүлэн авч үзнэ.",
    PART_3: "Энэ хэсэгт та хоёрын сэтгэлгээний онцлог, амьдралыг үзэх үзэл болон оюун санааны нийцлийг 'Мэнгэ'-ний зурхайгаар тайлбарлая.",
    PART_4: "Энэ хэсэг нь та хоёрын өдөр тутмын ахуй амьдрал, гэр орны уур амьсгал, санхүүгийн удирдлага болон гал тогоонд хэрхэн зохицохыг харуулна.",
    PART_5: "Энэ хэсэгт бид өмнөх дөрвөн хэсгийн мэдээллийг базаж, та хоёрын ирээдүйн ерөнхий зураглал болон үйлийн үрийн холбоог эцэслэн дүгнэж байна."
  }
};

// --- MAIN FUNCTION ---
function main() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const rows = sheet.getDataRange().getValues();
  const KEYS = {
    GEMINI: PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY"),
    MANYCHAT: PropertiesService.getScriptProperties().getProperty("MANYCHAT_API_TOKEN"),
    TEMPLATE: PropertiesService.getScriptProperties().getProperty("TEMPLATE_ID") 
  };

  let processedCount = 0;
  const START_TIME = new Date().getTime();

  try {
    for (let i = 1; i < rows.length; i++) {
      if (new Date().getTime() - START_TIME > CONFIG.MAX_EXECUTION_TIME - CONFIG.SAFETY_BUFFER) break;
      if (processedCount >= CONFIG.BATCH_SIZE) break;

      const row = rows[i];
      if (row[CONFIG.COLUMNS.STATUS] === "DONE" || String(row[CONFIG.COLUMNS.STATUS]).includes("ERROR") || !row[CONFIG.COLUMNS.INPUT]) continue;

      sheet.getRange(i + 1, CONFIG.COLUMNS.STATUS + 1).setValue("Processing...");
      SpreadsheetApp.flush();

      try {
        const inputString = String(row[CONFIG.COLUMNS.INPUT]); 
        
        // 1. DATA PREP: AI PARSING
        const normalized = normalizeInputWithAI(inputString, KEYS.GEMINI);
        
        // 2. PARSE & CALCULATE (The Brain)
        const coupleData = calculateCoupleProfile(normalized);
        
        // 3. GENERATE REPORT (The Voice) - 6 STEP CALLS + HEADER
        const reportResult = generateSequentialReport(coupleData, KEYS.GEMINI);
        
        // 4. DELIVERY
        const pdfUrl = createPdf(`${coupleData.user.name} & ${coupleData.partner.name}`, reportResult.text, KEYS.TEMPLATE);
        sendManyChat(row[CONFIG.COLUMNS.ID], pdfUrl, coupleData.user.name || "Erhem", KEYS.MANYCHAT);

        // LOGGING
        const totalTokens = (normalized.usage || 0) + reportResult.usage;
        const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
        sheet.getRange(i + 1, CONFIG.COLUMNS.PDF + 1).setValue(pdfUrl);
        sheet.getRange(i + 1, CONFIG.COLUMNS.STATUS + 1).setValue("DONE");
        sheet.getRange(i + 1, CONFIG.COLUMNS.TOKEN + 1).setValue(totalTokens); 
        sheet.getRange(i + 1, CONFIG.COLUMNS.DEBUG + 1).setValue(JSON.stringify(coupleData));
        sheet.getRange(i + 1, CONFIG.COLUMNS.DATE + 1).setValue(now);
        sheet.getRange(i + 1, CONFIG.COLUMNS.VER + 1).setValue(CONFIG.VERSION);
        sheet.getRange(i + 1, CONFIG.COLUMNS.ERROR + 1).setValue(""); 
        
        processedCount++;
      } catch (err) {
        console.error(err);
        sheet.getRange(i + 1, CONFIG.COLUMNS.STATUS + 1).setValue("ERROR");
        sheet.getRange(i + 1, CONFIG.COLUMNS.ERROR + 1).setValue(err.message);
      }
    }
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 1. AI NORMALIZATION
// ==========================================
function normalizeInputWithAI(raw, key) {
  const prompt = `
  TASK: Normalize this couple's data.
  INPUT: "${raw}"
  
  INSTRUCTIONS:
  1. Extract User's Gender (M/F).
  2. Extract User's DOB (YYYY.MM.DD).
  3. Extract Partner's DOB (YYYY.MM.DD).
  4. Use "Name" if present, else "User".
  
  RETURN JSON ONLY:
  {
    "userGender": "Эрэгтэй" or "Эмэгтэй",
    "dob1": "YYYY.MM.DD",
    "dob2": "YYYY.MM.DD",
    "name": "User"
  }
  `;
  
  try {
    const result = callGemini(prompt, key);
    const cleanJson = result.text.replace(/```json/g, "").replace(/```/g, "").trim();
    const data = JSON.parse(cleanJson);
    if (!data.dob1 || !data.dob2) throw new Error("AI Missing Fields");
    return { ...data, usage: result.usage };
  } catch (e) {
    console.warn("AI Normalization Failed, using Fallback Regex", e);
    const dates = raw.match(/\d{4}[\.\-\/]\d{2}[\.\-\/]\d{2}/g) || ["1990.01.01", "1990.01.01"];
    let gender = "Эмэгтэй";
    if (raw.toLowerCase().includes("male") || raw.toLowerCase().includes("эр")) gender = "Эрэгтэй";
    
    return {
      userGender: gender,
      dob1: dates[0] || "1990.01.01",
      dob2: dates[1] || "1992.02.02",
      name: "User",
      usage: 0
    };
  }
}

// ==========================================
// 2. CALCULATIONS ENGINE (Tsagaan Sar + Year Menge + Life Gua)
// ==========================================

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
  const animal = CONFIG.ANIMAL_MAP[(y - 1900) % 12];
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
  let k = (gender === "Эрэгтэй") ? 11 - n : n + 4;
  if (k > 9) k = sumDigits(k);
  
  let suudal = (k === 5) ? ((gender === "Эрэгтэй") ? "Хон" : "Гэн") : CONFIG.SUUDAL_NAMES[k];
  const suudalElement = CONFIG.SUUDAL_ELEMENTS[suudal];
  
  return { dob, year: y, gender, name, animal, element, menge: mengeNum, mengeName, mengeElement, suudal, suudalElement };
}

function sumDigits(num) {
  return String(num).split('').reduce((a, b) => a + Number(b), 0);
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

// ==========================================
// 3. GENERATION ENGINE (MAX SPLIT + 8192 TOKENS + STRUCTURED PROMPTS)
// ==========================================

function generateSequentialReport(data, apiKey) {
  const { user, partner, relations } = data;
  
  let male = (user.gender === "Эрэгтэй") ? user : partner;
  let female = (user.gender === "Эмэгтэй") ? user : partner;

  // --- HEADER GENERATION (HARD-CODED) ---
  const header = `
📊 ХЭРЭГЛЭГЧИЙН МЭДЭЭЛЭЛ

👨 **ЭРЭГТЭЙН ӨГӨГДӨЛ**
📅 Огноо: ${male.dob}
🐅 Жил: ${male.element} ${male.animal}
⚪ Мэнгэ: ${male.mengeName}
🏔️ Суудал: ${male.suudal}
🔥 Махбод: ${male.element}

👩 **ЭМЭГТЭЙН ӨГӨГДӨЛ**
📅 Огноо: ${female.dob}
🐇 Жил: ${female.element} ${female.animal}
⚪ Мэнгэ: ${female.mengeName}
🏔️ Суудал: ${female.suudal}
⛰️ Махбод: ${female.element}

---
`;

  // STRONG SYSTEM PROMPT (Foreign Words Forbidden)
  const SYSTEM_PROMPT = `
  ROLE: Master Mongolian Astrologer.
  GOAL: Write a detailed, philosophical report based on the "Gold Standard" style.
  RULES:
  1. DO NOT STOP after the intro. You must write the FULL ANALYSIS.
  2. Use the REFERENCE TEXT as a STYLE GUIDE. Do not copy it blindly.
  3. NO FOREIGN WORDS (Korean, English, etc). Use PURE MONGOLIAN.
  4. DO NOT use "---" separator lines.
  5. Use EMOJIS instead of "*" for list items.
  `;

  // --- CALL 1: PART 1 ONLY (Year) ---
  const promptPart1 = `
  ${SYSTEM_PROMPT}
  TASK: Write PART 1 ONLY.
  
  **REQUIRED STRUCTURE:**
  STEP 1: Write Title: "**1-р хэсэг. ЯЗГУУР МӨН ЧАНАР: Жилүүдийн Ивээл ба Харш**"
  STEP 2: Write Intro: "${CONFIG.INTRO_TEXTS.PART_1}"
  STEP 3: Write CHARACTER ANALYSIS for Husband (${male.year} ${male.element} ${male.animal}).
  STEP 4: Write CHARACTER ANALYSIS for Wife (${female.year} ${female.element} ${female.animal}).
  STEP 5: Write COMPATIBILITY LOGIC based on Relation: "${relations.element}". Explain *why* it is Good/Bad.
  STEP 6: Write Conclusion for Part 1.
  
  DATA:
  - Husband: ${male.animal}, Element: ${male.element}
  - Wife: ${female.animal}, Element: ${female.element}
  - Relation: ${relations.element}
  
  REFERENCE STYLE PART 1 (Use this tone):
  ${CONFIG.REFERENCES.PART_1}
  `;
  const call1 = callGemini(promptPart1, apiKey);

  // --- CALL 2: PART 2 ONLY (Energy) ---
  const promptPart2 = `
  ${SYSTEM_PROMPT}
  TASK: Write PART 2 ONLY.
  
  **REQUIRED STRUCTURE:**
  STEP 1: Write Title: "**2-р хэсэг. ЭРЧИМ ХҮЧНИЙ НИЙЛЭМЖ: Махбодын Зохицол**"
  STEP 2: Write Intro: "${CONFIG.INTRO_TEXTS.PART_2}"
  STEP 3: Explain ENERGY FLOW (${male.element} vs ${female.element}). Who feeds whom? Or are they same?
  STEP 4: Explain EMOTIONAL RHYTHM & HEALTH based on elements.
  STEP 5: Write Conclusion for Part 2.
  
  DATA:
  - Husband: ${male.animal}, Element: ${male.element}
  - Wife: ${female.animal}, Element: ${female.element}
  - Relation: ${relations.element}

  REFERENCE STYLE PART 2 (Use this tone):
  ${CONFIG.REFERENCES.PART_2}
  `;
  const call2 = callGemini(promptPart2, apiKey);

  // --- CALL 3: PART 3 ONLY (Menge) ---
  const promptPart3 = `
  ${SYSTEM_PROMPT}
  TASK: Write PART 3 ONLY.
  
  **REQUIRED STRUCTURE:**
  STEP 1: Write Title: "**3-р хэсэг. ОЮУН САНААНЫ ХОЛБООС: Мэнгэний Таарамж**"
  STEP 2: Write Intro: "${CONFIG.INTRO_TEXTS.PART_3}"
  STEP 3: Analyze MINDSET of Male (${male.mengeName}).
  STEP 4: Analyze MINDSET of Female (${female.mengeName}).
  STEP 5: Explain CONNECTION "${relations.menge}".
  
  DATA:
  - Male Menge: ${male.mengeName} (${male.mengeElement})
  - Female Menge: ${female.mengeName} (${female.mengeElement})
  - Menge Relation: ${relations.menge}
  
  REFERENCE STYLE PART 3 (Use this tone):
  ${CONFIG.REFERENCES.PART_3}
  `;
  const call3 = callGemini(promptPart3, apiKey);

  // --- CALL 4: PART 4 ONLY (Suudal) ---
  const promptPart4 = `
  ${SYSTEM_PROMPT}
  TASK: Write PART 4 ONLY.
  
  **REQUIRED STRUCTURE:**
  STEP 1: Write Title: "**4-р хэсэг. АМЬДРАЛЫН ХЭМНЭЛ: Суудлын Өнцөг & Гал Голомт**"
  STEP 2: Write Intro: "${CONFIG.INTRO_TEXTS.PART_4}"
  STEP 3: Analyze LIFESTYLE of Male (${male.suudal}) vs Female (${female.suudal}).
  STEP 4: Explain CONFLICT "${relations.suudal}".
  STEP 5: EXPLAIN REMEDY: Use "${relations.remedy.element}" to bridge gap. Advice: "${relations.remedy.advice}".
  
  DATA:
  - Male Suudal: ${male.suudal} (${male.suudalElement})
  - Female Suudal: ${female.suudal} (${female.suudalElement})
  - Suudal Relation: ${relations.suudal}
  - Remedy: ${relations.remedy.element}
  
  REFERENCE STYLE PART 4 (Use this tone):
  ${CONFIG.REFERENCES.PART_4}
  `;
  const call4 = callGemini(promptPart4, apiKey);

  // --- CALL 5: PART 5 ONLY (Conclusion) ---
  const prompt5 = `
  ${SYSTEM_PROMPT}
  TASK: Write PART 5 (Conclusion).
  
  **REQUIRED STRUCTURE:**
  STEP 1: Write the Title: "**5-р хэсэг. ҮЙЛИЙН ҮР БА ИРЭЭДҮЙН ЗУРАГЛАЛ Нэгдсэн Дүгнэлт**"
  STEP 2: Write this exact Intro: "${CONFIG.INTRO_TEXTS.PART_5}"
  STEP 3: Explain KARMIC CONNECTION. Why did they meet?
  STEP 4: Future Outlook & Happiness Key (Remedy: ${relations.remedy.element}).
  
  DATA:
  - Score: ${relations.score}%
  - Remedy: ${relations.remedy.element}
  
  STYLE GUIDE (REFERENCE):
  ${CONFIG.REFERENCES.PART_5}
  `;
  const call5 = callGemini(prompt5, apiKey);
  
  return {
    text: [header, call1.text, call2.text, call3.text, call4.text, call5.text].join("\n\n"),
    usage: call1.usage + call2.usage + call3.usage + call4.usage + call5.usage
  };
}

// --- UTILS ---
function callGemini(text, key, attempts = 2) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.GEMINI_MODEL}:generateContent?key=${key}`;
  const payload = {
    contents: [{ parts: [{ text: text }] }],
    generationConfig: { 
      temperature: CONFIG.TEMPERATURE, 
      maxOutputTokens: 8192 // INCREASED TO 8192 TO PREVENT CUTOFF
    }
  };
  const options = { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true };
  
  for (let i = 0; i < attempts; i++) {
    try {
      const res = UrlFetchApp.fetch(url, options);
      const json = JSON.parse(res.getContentText());
      if (json.candidates) return { text: json.candidates[0].content.parts[0].text.trim(), usage: json.usageMetadata ? json.usageMetadata.totalTokenCount : 0 };
    } catch (e) {
      if (i === attempts - 1) return { text: "Error generating content.", usage: 0 };
      Utilities.sleep(1000);
    }
  }
}

function createPdf(name, content, templateId) {
  const folderId = PropertiesService.getScriptProperties().getProperty("FOLDER_ID");
  const copy = DriveApp.getFileById(templateId).makeCopy(`${name} - Harmony Report`);
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();
  
  body.replaceText("{{NAME}}", name);
  
  // 3. CLEAN ARTIFACTS (STRICT)
  // Remove "---" or "***" separators
  const cleanContent = content
    .replace(/\n\s*---\s*\n/g, "\n")
    .replace(/\n\s*\*\*\*\s*\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n"); // Normalize max 2 newlines
    
  body.replaceText("{{REPORT}}", cleanContent);
  
  // FIXED FORMATTING ENGINE
  processFormatting(body);

  // Paragraph Styling
  var paras = body.getParagraphs();
  for (var i = 0; i < paras.length; i++) {
    var p = paras[i];
    p.setLineSpacing(1.15);
    p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
  }
  
  doc.saveAndClose();
  const pdf = copy.getAs(MimeType.PDF);
  const file = DriveApp.getFolderById(folderId).createFile(pdf);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  copy.setTrashed(true);
  return file.getUrl();
}

// FORMATTING ENGINE
function processFormatting(body) {
  // 1. Handle BOLD: **text** -> Bold Text
  var foundElement = body.findText("\\*\\*([^*\\n]+?)\\*\\*");
  while (foundElement != null) {
    var foundText = foundElement.getElement().asText();
    var start = foundElement.getStartOffset();
    var end = foundElement.getEndOffsetInclusive();
    foundText.setBold(start, end, true);
    foundText.deleteText(end - 1, end); 
    foundText.deleteText(start, start + 1); 
    foundElement = body.findText("\\*\\*([^*\\n]+?)\\*\\*");
  }
  
  // 2. Handle BULLETS: Remove "* " or "- " if used, relying on Emojis
  var paras = body.getParagraphs();
  for (var i = 0; i < paras.length; i++) {
    var p = paras[i];
    var text = p.getText();
    if ((text.trim().startsWith("* ") || text.trim().startsWith("- ")) && !text.trim().startsWith("**")) {
      p.setText(text.replace(/^\s*[\*\-]\s+/, "")); 
    }
  }
}

function sendManyChat(subscriberId, pdfUrl, name, token) {
  const msg = CONFIG.DELIVERY_MESSAGE.replace("{{URL}}", pdfUrl);
  const url = "https://api.manychat.com/fb/sending/sendContent";
  const payload = { "subscriber_id": String(subscriberId).trim(), data: { version: "v2", content: { messages: [{ type: "text", text: msg }] } } };
  UrlFetchApp.fetch(url, { method: "post", headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" }, payload: JSON.stringify(payload), muteHttpExceptions: true });
}
