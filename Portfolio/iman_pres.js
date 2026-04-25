const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.title = 'أركان الإيمان';

// === COLORS ===
const GD = "1A3C28"; // dark green
const GM = "2D6A4F"; // medium green
const GL = "4A8C65"; // light green
const GO = "C9A227"; // gold
const GX = "E8CC6A"; // gold light
const CR = "FBF3E2"; // cream
const WH = "FFFFFF"; // white
const TX = "1A3C28"; // text dark

// === HELPERS ===
function frame(s) {
  const m = 0.18, t = 0.032;
  const rects = [
    {x:m, y:m, w:10-2*m, h:t},
    {x:m, y:5.625-m-t, w:10-2*m, h:t},
    {x:m, y:m, w:t, h:5.625-2*m},
    {x:10-m-t, y:m, w:t, h:5.625-2*m}
  ];
  rects.forEach(r => s.addShape(pres.shapes.RECTANGLE, {...r, fill:{color:GO}, line:{width:0}}));
  [[0.12,0.12],[9.78,0.12],[0.12,5.4],[9.78,5.4]].forEach(([x,y]) =>
    s.addShape(pres.shapes.RECTANGLE, {x, y, w:0.1, h:0.1, rotate:45, fill:{color:GO}, line:{width:0}})
  );
}

function topBar(s, title, bg=GD, textColor=GO, fontSize=26) {
  s.addShape(pres.shapes.RECTANGLE, {x:0, y:0, w:10, h:1.0, fill:{color:bg}, line:{width:0}});
  // Gold accent strip
  s.addShape(pres.shapes.RECTANGLE, {x:0, y:0.96, w:10, h:0.04, fill:{color:GO}, line:{width:0}});
  s.addText(title, {x:0.4, y:0.1, w:9.2, h:0.78, fontSize, bold:true, color:textColor, align:"center", rtlMode:true, fontFace:"Arial"});
}

function div(s, y, x=0.5, w=9) {
  s.addShape(pres.shapes.RECTANGLE, {x, y, w, h:0.03, fill:{color:GO}, line:{width:0}});
  // Small diamond in center
  s.addShape(pres.shapes.RECTANGLE, {x:4.93, y:y-0.05, w:0.1, h:0.1, rotate:45, fill:{color:GO}, line:{width:0}});
}

function crd(s, x, y, w, h, fill, text, tColor=WH, fs=16, border=null, bold=false) {
  s.addShape(pres.shapes.RECTANGLE, {x, y, w, h, fill:{color:fill},
    line: border ? {color:border, width:2} : {width:0},
    shadow:{type:"outer", blur:5, offset:2, angle:135, color:"000000", opacity:0.18}
  });
  if(text) s.addText(text, {x:x+0.1, y, w:w-0.2, h, fontSize:fs, bold, color:tColor, align:"center", valign:"middle", rtlMode:true, fontFace:"Arial"});
}

function bul(s, items, x, y, w, h, fontSize=15, color=TX, spacing=0.62) {
  items.forEach((item, i) => {
    // Gold diamond bullet
    s.addShape(pres.shapes.RECTANGLE, {x:x+w-0.22, y:y+i*spacing+0.14, w:0.09, h:0.09, rotate:45, fill:{color:GO}, line:{width:0}});
    s.addText(item, {x, y:y+i*spacing, w:w-0.3, h:spacing-0.05, fontSize, color, align:"right", valign:"middle", rtlMode:true, fontFace:"Arial"});
  });
}

function numberedCards(s, items, startX, startY, cardW, cardH, gap, cols=2) {
  items.forEach((item, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = startX + col * (cardW + gap);
    const y = startY + row * (cardH + gap);
    // Background card
    s.addShape(pres.shapes.RECTANGLE, {x, y, w:cardW, h:cardH, fill:{color:GM},
      line:{color:GO, width:1.5},
      shadow:{type:"outer", blur:4, offset:2, angle:135, color:"000000", opacity:0.2}
    });
    // Number circle
    s.addShape(pres.shapes.OVAL, {x:x+cardW-0.5, y:y+0.06, w:0.38, h:0.38, fill:{color:GO}, line:{width:0}});
    s.addText(String(i+1), {x:x+cardW-0.5, y:y+0.06, w:0.38, h:0.38, fontSize:14, bold:true, color:GD, align:"center", valign:"middle"});
    s.addText(item, {x:x+0.1, y, w:cardW-0.65, h:cardH, fontSize:15, color:WH, align:"right", valign:"middle", rtlMode:true, fontFace:"Arial"});
  });
}

function ornament(s, x, y, size=28) {
  s.addText("❋", {x, y, w:size/96+0.3, h:size/96+0.1, fontSize:size, color:GO, align:"center"});
}

// ==================================================
// SLIDE 1 - TITLE
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: GD};
  frame(s);
  // Ornamental stars
  s.addText("✦  ❋  ✦", {x:0.5, y:0.5, w:9, h:0.55, fontSize:24, color:GO, align:"center"});
  s.addText("أركانُ الإيمانِ", {
    x:0.5, y:1.1, w:9, h:1.2,
    fontSize:56, bold:true, color:GO, align:"center", rtlMode:true, fontFace:"Arial"
  });
  s.addText("الرُّكنُ الأوَّل: الإيمانُ باللهِ", {
    x:0.5, y:2.35, w:9, h:0.75,
    fontSize:30, color:WH, align:"center", rtlMode:true, fontFace:"Arial"
  });
  div(s, 3.22, 2.5, 5);
  s.addText("الصف الثالث الإعدادي", {
    x:0.5, y:3.35, w:9, h:0.42, fontSize:18, color:GX, align:"center", rtlMode:true
  });
  s.addText("البناء العقدي للجيل الصاعد | صفحات ٤٩ - ٧١", {
    x:0.5, y:3.82, w:9, h:0.35, fontSize:14, color:GX, align:"center", rtlMode:true
  });
  s.addText("﴿ آمَنَ الرَّسُولُ بِمَا أُنزِلَ إِلَيْهِ مِن رَّبِّهِ وَالْمُؤْمِنُونَ ﴾", {
    x:0.5, y:4.78, w:9, h:0.42, fontSize:14, color:GX, italic:true, align:"center", rtlMode:true
  });
  s.addText("✦  ❋  ✦", {x:0.5, y:5.15, w:9, h:0.3, fontSize:14, color:GO, align:"center"});
}

// ==================================================
// SLIDE 2 - OBJECTIVES
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: CR};
  frame(s);
  topBar(s, "أهداف الدرس");
  s.addText("في نهاية هذه المحاضرة سيكون الطالب قادراً على:", {
    x:0.5, y:1.08, w:9, h:0.45, fontSize:15, color:GM, align:"right", rtlMode:true, italic:true
  });
  const objectives = [
    "الاستدلال بالعقل والفطرة على وجود الله ووحدانيته",
    "التعرف على أهمية أسماء الله الحسنى وكيفية التعامل معها",
    "التمييز بين توحيد الربوبية وتوحيد الألوهية",
    "إدراك ثمرات التوحيد وأثره في الحياة اليومية",
    "التعرف على مظاهر ضعف العقيدة وكيفية تجنبها",
  ];
  objectives.forEach((obj, i) => {
    const y = 1.6 + i * 0.78;
    s.addShape(pres.shapes.OVAL, {x:8.7, y:y+0.08, w:0.38, h:0.38, fill:{color:GM}, line:{width:0}});
    s.addText(String(i+1), {x:8.7, y:y+0.08, w:0.38, h:0.38, fontSize:14, bold:true, color:WH, align:"center", valign:"middle"});
    s.addShape(pres.shapes.RECTANGLE, {x:0.45, y:y, w:8.15, h:0.6, fill:{color:"EEF6F1"}, line:{color:GM, width:0.7}});
    s.addText(obj, {x:0.55, y, w:7.95, h:0.6, fontSize:15, color:TX, align:"right", valign:"middle", rtlMode:true, fontFace:"Arial"});
  });
}

// ==================================================
// SLIDE 3 - SIX PILLARS OVERVIEW
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: GD};
  frame(s);
  ornament(s, 4.7, 0.25, 24);
  s.addText("أركانُ الإيمانِ السِّتَّة", {
    x:0.5, y:0.55, w:9, h:0.7, fontSize:34, bold:true, color:GO, align:"center", rtlMode:true
  });
  s.addText("قال النبي ﷺ: «الإيمانُ أَن تُؤمِنَ بِاللهِ وَمَلائِكَتِهِ وَكُتُبِهِ وَرُسُلِهِ وَاليَومِ الآخِرِ وَالقَدَرِ خَيرِهِ وَشَرِّهِ»", {
    x:0.5, y:1.25, w:9, h:0.65, fontSize:14, color:GX, italic:true, align:"center", rtlMode:true
  });
  div(s, 1.95, 1.5, 7);
  const pillars = [
    "الإيمان بالله", "الإيمان بالملائكة", "الإيمان بالكتب",
    "الإيمان بالرُّسُل", "الإيمان باليوم الآخر", "الإيمان بالقَدَر"
  ];
  pillars.forEach((p, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = col === 0 ? 5.1 : 0.4;
    const y = 2.1 + row * 1.05;
    const isFirst = i === 0;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w:4.45, h:0.85,
      fill:{color: isFirst ? GO : GM},
      line:{color:GO, width: isFirst ? 0 : 1.5},
      shadow:{type:"outer", blur:5, offset:2, angle:135, color:"000000", opacity:0.25}
    });
    s.addShape(pres.shapes.RECTANGLE, {x: col===0 ? x : x+3.95, y, w:0.5, h:0.85, fill:{color: isFirst ? GD : GD}, line:{width:0}});
    s.addText(String(i+1), {x: col===0 ? x : x+3.95, y, w:0.5, h:0.85, fontSize:18, bold:true, color: isFirst ? GO : GO, align:"center", valign:"middle"});
    s.addText(p + (isFirst ? " ◄" : ""), {
      x: col===0 ? x+0.55 : x+0.1, y, w:3.75, h:0.85,
      fontSize: isFirst ? 17 : 16, bold: isFirst, color: isFirst ? GD : WH,
      align:"right", valign:"middle", rtlMode:true, fontFace:"Arial"
    });
  });
}

// ==================================================
// SLIDE 4 - FIRST PILLAR INTRO
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: CR};
  frame(s);
  topBar(s, "الرُّكنُ الأوَّل: الإيمانُ باللهِ");
  s.addText("أَصلُ الأُصُول", {
    x:0.5, y:1.1, w:9, h:0.55, fontSize:26, bold:true, color:GO, align:"center", rtlMode:true
  });
  div(s, 1.72, 2, 6);
  s.addText("الإيمانُ باللهِ هو أصلُ الأُصُول، وما يليه من الأركان إنما هو فرعٌ عنه؛ فاللهُ هو الذي أخبرنا بالملائكة والرسل واليوم الآخر والقدر، فبدون الإيمان به لا يمكن الإيمان ببقية الأركان", {
    x:0.6, y:1.85, w:8.8, h:1.2, fontSize:17, color:TX, align:"right", rtlMode:true, fontFace:"Arial"
  });
  div(s, 3.15, 1.5, 7);
  s.addText("يشتملُ الإيمانُ باللهِ على ثلاثة محاور:", {
    x:0.5, y:3.25, w:9, h:0.45, fontSize:17, bold:true, color:GM, align:"right", rtlMode:true
  });
  const topics = ["وُجُودُ اللهِ وكَمَالُه ورُبُوبِيَّتُه", "أَسمَاءُ اللهِ وصِفَاتُه", "أُلُوهِيَّتُه سُبحَانَه"];
  topics.forEach((t, i) => {
    crd(s, 0.4 + i*3.15, 3.78, 2.95, 1.3, GM, t, WH, 16, GO, true);
    s.addShape(pres.shapes.RECTANGLE, {x:0.4+i*3.15, y:3.78, w:2.95, h:0.08, fill:{color:GO}, line:{width:0}});
    s.addText(["١", "٢", "٣"][i], {x:0.4+i*3.15, y:3.82, w:0.5, h:0.5, fontSize:16, bold:true, color:WH, align:"center", valign:"middle"});
  });
}

// ==================================================
// SLIDE 5 - SECTION HEADER: وجود الله
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: GD};
  frame(s);
  s.addText("❋", {x:4.5, y:0.8, w:1, h:0.7, fontSize:36, color:GO, align:"center"});
  s.addText("المحورُ الأوَّل", {
    x:0.5, y:1.5, w:9, h:0.6, fontSize:22, color:GX, align:"center", rtlMode:true
  });
  s.addText("وُجُودُ اللهِ وكَمَالُه ورُبُوبِيَّتُه", {
    x:0.5, y:2.1, w:9, h:1.0, fontSize:42, bold:true, color:GO, align:"center", rtlMode:true, fontFace:"Arial"
  });
  div(s, 3.2, 2, 6);
  s.addText("كيف نعرف أن الله موجود؟", {
    x:0.5, y:3.35, w:9, h:0.55, fontSize:22, color:WH, align:"center", rtlMode:true, italic:true
  });
  s.addText("دليلان كبيران: العَقلُ والفِطرَة", {
    x:0.5, y:3.95, w:9, h:0.5, fontSize:20, color:GX, align:"center", rtlMode:true
  });
  s.addText("✦  ✦  ✦", {x:0.5, y:4.75, w:9, h:0.4, fontSize:18, color:GO, align:"center"});
}

// ==================================================
// SLIDE 6 - RATIONAL ARGUMENT
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: CR};
  frame(s);
  topBar(s, "أولاً: دليلُ العَقلِ على وُجودِ الله");
  s.addText("المبدأ الأساسي:", {x:0.5, y:1.1, w:9, h:0.42, fontSize:17, bold:true, color:GO, align:"right", rtlMode:true});
  
  // Central principle card
  s.addShape(pres.shapes.RECTANGLE, {x:0.5, y:1.55, w:9, h:1.05,
    fill:{color:GD}, line:{color:GO, width:2},
    shadow:{type:"outer", blur:6, offset:2, angle:135, color:"000000", opacity:0.25}});
  s.addText("الاستدلالُ بالأثرِ على المُؤثِّر", {
    x:0.5, y:1.55, w:9, h:1.05, fontSize:28, bold:true, color:GO, align:"center", valign:"middle", rtlMode:true
  });
  
  div(s, 2.72);
  s.addText("هذا المبدأ فطري يجده الإنسان في عقله بشكل طبيعي:", {
    x:0.5, y:2.82, w:9, h:0.42, fontSize:15, color:TX, align:"right", rtlMode:true, italic:true
  });
  
  const examples = [
    "إذا رأيتَ بيتاً → عرفتَ أن له بانياً",
    "إذا رأيتَ سيارةً → عرفتَ أن لها صانعاً",
    "إذا رأيتَ الكونَ المتقن → فلا بد له من خالقٍ عظيم",
  ];
  examples.forEach((e, i) => {
    const y = 3.3 + i * 0.67;
    s.addShape(pres.shapes.RECTANGLE, {x:0.45, y, w:9.1, h:0.55, fill:{color:"EEF6F1"}, line:{color:GL, width:0.8}});
    s.addShape(pres.shapes.RECTANGLE, {x:9.15, y, w:0.4, h:0.55, fill:{color:GM}, line:{width:0}});
    s.addText(["أ","ب","ج"][i], {x:9.15, y, w:0.4, h:0.55, fontSize:15, bold:true, color:WH, align:"center", valign:"middle"});
    s.addText(e, {x:0.55, y, w:8.5, h:0.55, fontSize:15, color:TX, align:"right", valign:"middle", rtlMode:true, fontFace:"Arial"});
  });
}

// ==================================================
// SLIDE 7 - USB & DNA EXAMPLE
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: CR};
  frame(s);
  topBar(s, "مثالٌ مِن الحياةِ: الـUSB وجسمُ الإنسان");
  
  // USB side
  crd(s, 0.4, 1.1, 4.4, 3.95, GM, null, WH, 15, GO);
  s.addShape(pres.shapes.RECTANGLE, {x:0.4, y:1.1, w:4.4, h:0.55, fill:{color:GD}, line:{width:0}});
  s.addText("USB  💾", {x:0.4, y:1.1, w:4.4, h:0.55, fontSize:20, bold:true, color:GO, align:"center", valign:"middle"});
  const usbPoints = ["جهازٌ صغير الحجم","يخزّن بياناتٍ كثيرة","صنعه فريقٌ من العلماء","نُعجَب بصانعه ونُقِرّ بعلمه"];
  usbPoints.forEach((p, i) => {
    s.addShape(pres.shapes.RECTANGLE, {x:3.95, y:1.75+i*0.75, w:0.09, h:0.09, rotate:45, fill:{color:GO}, line:{width:0}});
    s.addText(p, {x:0.55, y:1.7+i*0.75, w:3.25, h:0.6, fontSize:15, color:WH, align:"right", valign:"middle", rtlMode:true});
  });
  
  // VS
  s.addShape(pres.shapes.OVAL, {x:4.53, y:2.7, w:0.94, h:0.94, fill:{color:GO}, line:{width:0}});
  s.addText("VS", {x:4.53, y:2.7, w:0.94, h:0.94, fontSize:18, bold:true, color:GD, align:"center", valign:"middle"});
  
  // Human DNA side
  crd(s, 5.2, 1.1, 4.4, 3.95, GD, null, WH, 15, GO);
  s.addShape(pres.shapes.RECTANGLE, {x:5.2, y:1.1, w:4.4, h:0.55, fill:{color:GM}, line:{width:0}});
  s.addText("جسمُ الإنسان 🧬", {x:5.2, y:1.1, w:4.4, h:0.55, fontSize:20, bold:true, color:GO, align:"center", valign:"middle"});
  const dnaPoints = ["أصغر بآلاف المرات من USB","يخزّن أضعافَ المعلومات","الـDNA وحده يحتوي معلوماتٍ هائلة","فكيف لا يكون له خالقٌ عظيم؟!"];
  dnaPoints.forEach((p, i) => {
    s.addShape(pres.shapes.RECTANGLE, {x:5.25, y:1.75+i*0.75, w:0.09, h:0.09, rotate:45, fill:{color:GO}, line:{width:0}});
    s.addText(p, {x:5.35, y:1.7+i*0.75, w:4.1, h:0.6, fontSize:15, color:WH, align:"right", valign:"middle", rtlMode:true});
  });
  
  // Bottom conclusion
  s.addShape(pres.shapes.RECTANGLE, {x:0.4, y:5.1, w:9.2, h:0.35, fill:{color:GO}, line:{width:0}});
  s.addText("كلّما كان الشيءُ أكثرَ إتقاناً، دلّ على صانعٍ أعظم → فالإنسانُ يدلُّ على اللهِ العظيم", {
    x:0.4, y:5.1, w:9.2, h:0.35, fontSize:13, bold:true, color:GD, align:"center", valign:"middle", rtlMode:true
  });
}

// ==================================================
// SLIDE 8 - ROBOT vs HUMAN
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: CR};
  frame(s);
  topBar(s, "مثالٌ آخر: الرجلُ الآلي والإنسان");
  
  s.addText("🤖", {x:0.8, y:1.15, w:1.5, h:1.5, fontSize:60, align:"center"});
  s.addText("الرجلُ الآلي (الروبوت)", {x:0.4, y:2.68, w:4, h:0.42, fontSize:16, bold:true, color:GM, align:"center", rtlMode:true});
  const robotPoints = ["يُحاكي بعضَ حركات الإنسان","يصنعه علماءُ متخصصون","يحتاج سنواتٍ من البحث والتطوير","كلّما ازدادت المحاكاة → ازداد الإعجابُ بصانعيه"];
  robotPoints.forEach((p, i) => {
    s.addShape(pres.shapes.RECTANGLE, {x:0.45, y:3.15+i*0.53, w:9.1, h:0.46, fill:{color:i%2===0?"EEF6F1":"F5F0E4"}, line:{color:GL, width:0.5}});
    s.addShape(pres.shapes.RECTANGLE, {x:9.15, y:3.15+i*0.53, w:0.4, h:0.46, fill:{color: i<2 ? GM : GD}, line:{width:0}});
    s.addText(p, {x:0.55, y:3.15+i*0.53, w:8.5, h:0.46, fontSize:14, color:TX, align:"right", valign:"middle", rtlMode:true});
  });
  
  s.addText("🧑", {x:7.7, y:1.15, w:1.5, h:1.5, fontSize:60, align:"center"});
  s.addText("الإنسانُ", {x:7.7, y:2.68, w:1.9, h:0.42, fontSize:16, bold:true, color:GD, align:"center", rtlMode:true});
  
  div(s, 2.72);
  s.addText("إذا كان الروبوتُ يحتاج كلَّ هذا الجهد → فمن الجنون أن يقال: إن الإنسانَ المُتقَن وُجِد بلا خالق!", {
    x:0.5, y:3.15, w:9, h:0.46, fontSize:14, bold:true, color:GM, align:"right", rtlMode:true
  });
}

// ==================================================
// SLIDE 9 - QURANIC VERSES
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: GD};
  frame(s);
  ornament(s, 4.6, 0.3, 26);
  s.addText("آياتٌ قرآنيةٌ تدلُّ على وُجودِ الله", {
    x:0.5, y:0.7, w:9, h:0.62, fontSize:26, bold:true, color:GO, align:"center", rtlMode:true
  });
  div(s, 1.38, 1.5, 7);
  
  const verses = [
    {text:"﴿ أَفَلَا يَنظُرُونَ إِلَى الْإِبِلِ كَيْفَ خُلِقَتْ ﴾", ref:"[الغاشية: ١٧]"},
    {text:"﴿ وَفِي أَنفُسِكُمْ أَفَلَا تُبْصِرُونَ ﴾", ref:"[الذاريات: ٢١]"},
    {text:"﴿ أَمْ خُلِقُوا مِنْ غَيْرِ شَيْءٍ أَمْ هُمُ الْخَالِقُونَ ﴾", ref:"[الطور: ٣٥]"},
  ];
  
  verses.forEach((v, i) => {
    s.addShape(pres.shapes.RECTANGLE, {
      x:0.5, y:1.55+i*1.28, w:9, h:1.12,
      fill:{color:GM}, line:{color:GO, width:1.5},
      shadow:{type:"outer", blur:4, offset:2, angle:135, color:"000000", opacity:0.2}
    });
    s.addShape(pres.shapes.RECTANGLE, {x:0.5, y:1.55+i*1.28, w:9, h:0.06, fill:{color:GO}, line:{width:0}});
    s.addText(v.text, {
      x:0.6, y:1.62+i*1.28, w:8.8, h:0.72, fontSize:20, color:GX, align:"center", valign:"middle", rtlMode:true, italic:true
    });
    s.addText(v.ref, {
      x:0.6, y:2.32+i*1.28, w:8.8, h:0.3, fontSize:13, color:GL, align:"center", rtlMode:true
    });
  });
}

// ==================================================
// SLIDE 10 - GROUP ACTIVITY 1
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: GM};
  frame(s);
  s.addShape(pres.shapes.RECTANGLE, {x:0, y:0, w:10, h:1.0, fill:{color:GD}, line:{width:0}});
  s.addShape(pres.shapes.RECTANGLE, {x:0, y:0.97, w:10, h:0.05, fill:{color:GO}, line:{width:0}});
  s.addText("🌿  نشاطٌ جماعي  🌿", {
    x:0.5, y:0.1, w:9, h:0.78, fontSize:28, bold:true, color:GO, align:"center", rtlMode:true
  });
  
  s.addText("التفكير في آيات الله في الكون", {
    x:0.5, y:1.12, w:9, h:0.6, fontSize:24, bold:true, color:WH, align:"center", rtlMode:true
  });
  div(s, 1.82, 1.5, 7);
  
  s.addText("⏱ الوقت: ٥ دقائق", {x:0.5, y:1.95, w:9, h:0.42, fontSize:16, color:GX, align:"center", rtlMode:true});
  
  s.addShape(pres.shapes.RECTANGLE, {x:0.5, y:2.45, w:9, h:2.6, fill:{color:GD}, line:{color:GO, width:2}});
  s.addText("التعليمات:", {x:0.6, y:2.55, w:8.8, h:0.4, fontSize:16, bold:true, color:GO, align:"right", rtlMode:true});
  const instructions = [
    "انقسموا إلى مجموعات من ٣-٤ طلاب",
    "كلّ مجموعة تذكر ٣ أشياء في الكون تدلّ على عظمة الله وإتقان خلقه",
    "مثال: السماء، النجوم، جسم الإنسان، الماء، النبات...",
    "قدِّموا إجاباتكم للصف",
  ];
  instructions.forEach((inst, i) => {
    s.addShape(pres.shapes.RECTANGLE, {x:9.05, y:3.0+i*0.55, w:0.09, h:0.09, rotate:45, fill:{color:GO}, line:{width:0}});
    s.addText(inst, {x:0.65, y:2.98+i*0.55, w:8.2, h:0.5, fontSize:14, color:WH, align:"right", valign:"middle", rtlMode:true});
  });
}

// ==================================================
// SLIDE 11 - FITRAH EVIDENCE INTRO
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: CR};
  frame(s);
  topBar(s, "ثانياً: دليلُ الفِطرةِ على وُجودِ اللهِ");
  
  s.addText("الفطرة: ما رَكَزَه اللهُ في الإنسان من الشعور بالافتقار إليه", {
    x:0.5, y:1.08, w:9, h:0.5, fontSize:16, color:GM, align:"right", rtlMode:true, italic:true
  });
  div(s, 1.65, 1.5, 7);
  
  s.addText("تدلُّ الفطرةُ على وُجودِ الخالق مِن ثلاثةِ جهات:", {
    x:0.5, y:1.78, w:9, h:0.48, fontSize:18, bold:true, color:TX, align:"right", rtlMode:true
  });
  
  const directions = [
    {num:"الجهة الأولى", title:"المعارف الأولية الفطرية", desc:"معارفٌ وُلِدت مع الإنسان، كمعرفة أن لكل حادثٍ محدِثاً"},
    {num:"الجهة الثانية", title:"الافتقار والتعبُّد", desc:"في فطرة الإنسان افتقارٌ ذاتي إلى قوة غيبية كاملة يتذلل لها"},
    {num:"الجهة الثالثة", title:"الغرائزُ والأخلاق", desc:"الغرائزُ الفطرية والقيمُ الأخلاقية لا تُفهَم إلا في ظلّ الإيمان بالخالق"},
  ];
  
  directions.forEach((d, i) => {
    const y = 2.38 + i * 1.0;
    s.addShape(pres.shapes.RECTANGLE, {x:0.45, y, w:9.1, h:0.85, fill:{color:GM}, line:{color:GO, width:1.5}, shadow:{type:"outer", blur:4, offset:1, angle:135, color:"000000", opacity:0.15}});
    s.addShape(pres.shapes.RECTANGLE, {x:0.45, y, w:1.6, h:0.85, fill:{color:GD}, line:{width:0}});
    s.addText(d.num, {x:0.45, y, w:1.6, h:0.85, fontSize:13, bold:true, color:GO, align:"center", valign:"middle", rtlMode:true});
    s.addText(d.title + " ←", {x:2.15, y, w:3.0, h:0.42, fontSize:15, bold:true, color:GX, align:"right", valign:"bottom", rtlMode:true});
    s.addText(d.desc, {x:2.15, y:y+0.42, w:7.2, h:0.43, fontSize:13, color:WH, align:"right", valign:"top", rtlMode:true});
  });
}

// ==================================================
// SLIDE 12 - FITRAH DIRECTION 1
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: CR};
  frame(s);
  topBar(s, "الجهةُ الأولى: المعارفُ الأوليةُ الفطريةُ");
  s.addText("هناك معارفُ وُلِدت مع الإنسان، لم يتعلمها في مدرسة ولم يتلقَّها في جامعة", {
    x:0.5, y:1.1, w:9, h:0.55, fontSize:15, color:TX, align:"right", rtlMode:true, italic:true
  });
  div(s, 1.73);
  s.addText("مثال:", {x:0.5, y:1.85, w:9, h:0.4, fontSize:16, bold:true, color:GO, align:"right", rtlMode:true});
  
  const examples = [
    {title:"لكلِّ حادثٍ محدِث", desc:"نعرف فطرياً أن أيَّ شيءٍ حدث لا بد له من سببٍ أوجده"},
    {title:"الجزءُ أصغرُ من الكل", desc:"معرفةٌ فطرية لا نحتاج دليلاً عليها"},
  ];
  examples.forEach((e, i) => {
    crd(s, 0.45, 2.32+i*1.55, 9.1, 1.25, i===0 ? GD : GM, null, WH, 15, GO);
    s.addShape(pres.shapes.RECTANGLE, {x:0.45, y:2.32+i*1.55, w:9.1, h:0.06, fill:{color:GO}, line:{width:0}});
    s.addText(e.title, {x:0.6, y:2.38+i*1.55, w:8.8, h:0.45, fontSize:20, bold:true, color:GO, align:"center", rtlMode:true});
    s.addText(e.desc, {x:0.6, y:2.85+i*1.55, w:8.8, h:0.65, fontSize:16, color:WH, align:"right", valign:"middle", rtlMode:true});
  });
  s.addText("→ هذه المعارف الفطرية تدلّنا على أن للكون خالقاً ومحدِثاً وهو الله", {
    x:0.5, y:5.1, w:9, h:0.38, fontSize:15, bold:true, color:GM, align:"right", rtlMode:true
  });
}

// ==================================================
// SLIDE 13 - FITRAH DIRECTION 2
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: CR};
  frame(s);
  topBar(s, "الجهةُ الثانية: الافتقارُ والتعبُّد");
  s.addText("في فطرة الإنسان افتقارٌ ذاتيٌّ إلى قوةٍ غيبيةٍ كاملة يرجو منها النفعَ ويستدفع بها الضرَّ", {
    x:0.5, y:1.1, w:9, h:0.6, fontSize:16, color:TX, align:"right", rtlMode:true
  });
  div(s, 1.77, 1.5, 7);
  
  s.addShape(pres.shapes.RECTANGLE, {x:0.45, y:1.92, w:9.1, h:1.45, fill:{color:GD}, line:{color:GO, width:2}});
  s.addText("الدليلُ التاريخي:", {x:0.55, y:1.98, w:8.9, h:0.42, fontSize:15, bold:true, color:GO, align:"right", rtlMode:true});
  s.addText("جميعُ الأمم في مختلف البلدان ومنذ قديم الزمان لها أماكن للعبادة، فحتى المشركون عبدوا الشمسَ والنارَ والأحجار، وهذا دليلٌ على أن الإنسان فطريٌّ على الحاجة إلى إله!", {
    x:0.55, y:2.42, w:8.9, h:0.88, fontSize:15, color:WH, align:"right", rtlMode:true
  });
  
  div(s, 3.45, 1.5, 7);
  s.addText("مهم:", {x:0.5, y:3.55, w:9, h:0.4, fontSize:16, bold:true, color:GM, align:"right", rtlMode:true});
  s.addText("الرسلُ حين بُعثوا إلى أقوامهم لم يكن محورُ رسالتهم إثباتَ وجود الخالق؛ لأن الأمم كانت تُقِرّ بذلك في الجملة، وإنما كان محورُ رسالتهم: الدعوةُ إلى إفراد الله بالعبادة", {
    x:0.5, y:4.0, w:9, h:1.1, fontSize:15, color:TX, align:"right", rtlMode:true
  });
}

// ==================================================
// SLIDE 14 - FITRAH DIRECTION 3
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: CR};
  frame(s);
  topBar(s, "الجهةُ الثالثة: الغرائزُ والأخلاقُ");
  s.addText("نرى في الإنسان والحيوان غرائزَ فطرية غير مكتسبة من المجتمع أو البيئة:", {
    x:0.5, y:1.1, w:9, h:0.5, fontSize:15, color:TX, align:"right", rtlMode:true
  });
  div(s, 1.67, 1.5, 7);
  
  const instincts = [
    {icon:"🍼", title:"الرضيع", desc:"يعرف كيف يرضع فور ولادته بدون تعليم"},
    {icon:"🦁", title:"الحيوان", desc:"يتجه إلى الرضاعة مباشرةً عند الولادة"},
    {icon:"💞", title:"الميل بين الجنسين", desc:"غريزةٌ فطرية في كل إنسان"},
    {icon:"⚖️", title:"القيم الأخلاقية", desc:"استحسانُ الصدق والعدل، واستقباحُ الظلم والقتل، فطرةٌ في كل إنسان"},
  ];
  instincts.forEach((inst, i) => {
    const col = i % 2, row = Math.floor(i/2);
    const x = col === 0 ? 5.0 : 0.4;
    const y = 1.85 + row * 1.52;
    s.addShape(pres.shapes.RECTANGLE, {x, y, w:4.4, h:1.35, fill:{color:GM}, line:{color:GO, width:1.5}});
    s.addText(inst.icon, {x, y:y+0.05, w:0.8, h:1.25, fontSize:26, align:"center", valign:"middle"});
    s.addText(inst.title, {x:x+0.85, y:y+0.08, w:3.45, h:0.4, fontSize:16, bold:true, color:GX, align:"right", rtlMode:true});
    s.addText(inst.desc, {x:x+0.85, y:y+0.5, w:3.45, h:0.75, fontSize:13, color:WH, align:"right", rtlMode:true});
  });
  
  s.addShape(pres.shapes.RECTANGLE, {x:0.4, y:5.0, w:9.2, h:0.44, fill:{color:GO}, line:{width:0}});
  s.addText("وجودُ الخير في الإنسان لا يُفهَم إلا في ظلِّ الإيمان بالخالق المُدبِّر الذي أَلهَم النفوسَ فجورَها وتقواها", {
    x:0.4, y:5.0, w:9.2, h:0.44, fontSize:13, bold:true, color:GD, align:"center", valign:"middle", rtlMode:true
  });
}

// ==================================================
// SLIDE 15 - SECTION 1 SUMMARY
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: GD};
  frame(s);
  ornament(s, 4.6, 0.28, 24);
  s.addText("خُلاصةُ المحورِ الأوَّل", {
    x:0.5, y:0.68, w:9, h:0.62, fontSize:30, bold:true, color:GO, align:"center", rtlMode:true
  });
  div(s, 1.38, 1.5, 7);
  
  const summaryPoints = [
    {icon:"🔍", text:"العقلُ يدلّنا على وجود الله من خلال مبدأ الاستدلال بالأثر على المؤثر"},
    {icon:"🌿", text:"الفطرةُ تدلّ على وجود الله من ثلاث جهات: المعارف الأولية، والافتقار، والغرائز والأخلاق"},
    {icon:"📖", text:"القرآنُ الكريم يدعو الإنسانَ إلى التأمل في آيات النفس والكون للتيقن بوجود الله"},
  ];
  summaryPoints.forEach((p, i) => {
    s.addShape(pres.shapes.RECTANGLE, {x:0.5, y:1.55+i*1.28, w:9, h:1.12, fill:{color:GM}, line:{color:GO, width:1.5}});
    s.addText(p.icon, {x:0.55, y:1.55+i*1.28, w:0.9, h:1.12, fontSize:30, align:"center", valign:"middle"});
    s.addText(p.text, {x:1.55, y:1.55+i*1.28, w:7.8, h:1.12, fontSize:16, color:WH, align:"right", valign:"middle", rtlMode:true});
  });
  
  s.addText("فَوَا عَجَباً كيف يُعصى الإلهُ  ***  أم كيف يجحده الجاحدُ", {
    x:0.5, y:5.05, w:9, h:0.42, fontSize:15, color:GX, italic:true, align:"center", rtlMode:true
  });
}

// ==================================================
// SLIDE 16 - SECTION HEADER: أسماء الله وصفاته
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: GD};
  frame(s);
  s.addText("❋", {x:4.5, y:0.8, w:1, h:0.7, fontSize:36, color:GO, align:"center"});
  s.addText("المحورُ الثاني", {x:0.5, y:1.52, w:9, h:0.58, fontSize:22, color:GX, align:"center", rtlMode:true});
  s.addText("أسماءُ اللهِ وصِفَاتُه", {
    x:0.5, y:2.1, w:9, h:1.0, fontSize:44, bold:true, color:GO, align:"center", rtlMode:true, fontFace:"Arial"
  });
  div(s, 3.2, 2, 6);
  s.addText("﴿ وَلِلَّهِ الْأَسْمَاءُ الْحُسْنَى فَادْعُوهُ بِهَا ﴾", {
    x:0.5, y:3.35, w:9, h:0.55, fontSize:22, color:GX, italic:true, align:"center", rtlMode:true
  });
  s.addText("[الأعراف: ١٨٠]", {x:0.5, y:3.95, w:9, h:0.35, fontSize:14, color:GL, align:"center", rtlMode:true});
  s.addText("✦  ✦  ✦", {x:0.5, y:4.75, w:9, h:0.4, fontSize:18, color:GO, align:"center"});
}

// ==================================================
// SLIDE 17 - IMPORTANCE OF NAMES
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: CR};
  frame(s);
  topBar(s, "أهميةُ التعرُّف على أسماءِ اللهِ الحُسنى");
  s.addText("الحديثُ عن أسماء الله وصفاته حديثٌ عظيمٌ جميل؛ لأنه يُعرِّفنا بالله ﷻ", {
    x:0.5, y:1.1, w:9, h:0.52, fontSize:15, color:TX, align:"right", rtlMode:true, italic:true
  });
  div(s, 1.7, 1.5, 7);
  const benefits = [
    {num:"١", title:"زيادة المحبة", desc:"كلما عرفتَ أسماءَ الله ازددتَ له حبّاً ومنه قُرباً"},
    {num:"٢", title:"إحسان الدعاء", desc:"ندعو اللهَ بأسمائه: «اللهمَّ يا رحيمُ ارحمني»"},
    {num:"٣", title:"التعبّد بها", desc:"نتأمل في معانيها ونتعبّد لله بمقتضاها"},
    {num:"٤", title:"الاقتداء بالأنبياء", desc:"الأنبياء أعرفُ الناس بالله، وكانوا يدعونه بأسمائه"},
  ];
  benefits.forEach((b, i) => {
    const col = i % 2, row = Math.floor(i/2);
    const x = col === 0 ? 5.05 : 0.45;
    const y = 1.88 + row * 1.65;
    s.addShape(pres.shapes.RECTANGLE, {x, y, w:4.4, h:1.45, fill:{color:GD}, line:{color:GO, width:2}, shadow:{type:"outer", blur:5, offset:2, angle:135, color:"000000", opacity:0.2}});
    s.addShape(pres.shapes.RECTANGLE, {x, y, w:4.4, h:0.5, fill:{color:GM}, line:{width:0}});
    s.addShape(pres.shapes.OVAL, {x:x+0.1, y:y+0.06, w:0.38, h:0.38, fill:{color:GO}, line:{width:0}});
    s.addText(b.num, {x:x+0.1, y:y+0.06, w:0.38, h:0.38, fontSize:14, bold:true, color:GD, align:"center", valign:"middle"});
    s.addText(b.title, {x:x+0.55, y, w:3.75, h:0.5, fontSize:16, bold:true, color:GX, align:"right", valign:"middle", rtlMode:true});
    s.addText(b.desc, {x:x+0.15, y:y+0.55, w:4.1, h:0.85, fontSize:14, color:WH, align:"right", valign:"middle", rtlMode:true});
  });
}

// ==================================================
// SLIDE 18 - HADITH OF 99 NAMES
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: GD};
  frame(s);
  ornament(s, 4.6, 0.28, 24);
  s.addText("حديثُ الأسماءِ الحُسنى", {
    x:0.5, y:0.68, w:9, h:0.62, fontSize:30, bold:true, color:GO, align:"center", rtlMode:true
  });
  div(s, 1.38, 1.5, 7);
  s.addShape(pres.shapes.RECTANGLE, {x:0.5, y:1.52, w:9, h:2.05, fill:{color:GM}, line:{color:GO, width:2}});
  s.addText("قال النبي ﷺ:", {x:0.65, y:1.62, w:8.7, h:0.45, fontSize:17, bold:true, color:GX, align:"right", rtlMode:true});
  s.addText("«إِنَّ لِلَّهِ تِسْعَةً وَتِسْعِينَ اسْماً، مِئَةً إِلَّا وَاحِداً، مَنْ أَحْصَاهَا دَخَلَ الْجَنَّةَ»", {
    x:0.65, y:2.1, w:8.7, h:1.32, fontSize:22, bold:true, color:WH, italic:true, align:"center", valign:"middle", rtlMode:true
  });
  s.addText("متفقٌ عليه | رواه البخاري ومسلم", {x:0.5, y:3.65, w:9, h:0.4, fontSize:14, color:GL, align:"center", rtlMode:true});
  div(s, 4.15, 1.5, 7);
  s.addText("٩٩ اسماً → مَن أحصاها دخل الجنة. وهذا لا يعني أن أسماء الله محصورة في ٩٩!", {
    x:0.5, y:4.28, w:9, h:0.6, fontSize:15, color:GX, align:"center", rtlMode:true
  });
  s.addText("«أسألُك بكلِّ اسمٍ هو لك... أو استأثَرتَ به في علمِ الغيب عندَك» [صحيح الألباني]", {
    x:0.5, y:4.92, w:9, h:0.38, fontSize:13, color:GL, italic:true, align:"center", rtlMode:true
  });
}

// ==================================================
// SLIDE 19 - MEANING OF IHSA'
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: CR};
  frame(s);
  topBar(s, "معنى «أَحْصَاهَا»");
  s.addText("اختلف العلماء في معنى «أَحصاها»، وأبرز الأقوال:", {
    x:0.5, y:1.1, w:9, h:0.45, fontSize:15, color:TX, align:"right", rtlMode:true
  });
  div(s, 1.62, 1.5, 7);
  
  const meanings = [
    {rank:"١", title:"الحفظُ والاستيعاب", desc:"قال البخاري وغيره: معناها حفظُها، وهو الأظهر لأنه جاء مفسَّراً في الرواية الأخرى: «مَن حفظَها دخل الجنة»", highlighted: true},
    {rank:"٢", title:"العدُّ والدعاء بها", desc:"أن يعدَّها في دعائه فيدعو الله بها واحداً واحداً"},
    {rank:"٣", title:"العمل بمقتضاها", desc:"أن يُحسِن المراعاةَ لها ويتحقق بمعانيها في سلوكه"},
  ];
  
  meanings.forEach((m, i) => {
    const y = 1.78 + i * 1.17;
    s.addShape(pres.shapes.RECTANGLE, {x:0.45, y, w:9.1, h:1.0, fill:{color: m.highlighted ? GD : GM}, line:{color:GO, width: m.highlighted ? 2 : 1}});
    s.addShape(pres.shapes.RECTANGLE, {x:0.45, y, w:1.2, h:1.0, fill:{color: m.highlighted ? GO : GL}, line:{width:0}});
    s.addText(m.rank, {x:0.45, y, w:1.2, h:1.0, fontSize:22, bold:true, color: m.highlighted ? GD : WH, align:"center", valign:"middle"});
    s.addText(m.title, {x:1.72, y:y+0.05, w:7.65, h:0.38, fontSize:17, bold:true, color: m.highlighted ? GX : WH, align:"right", rtlMode:true});
    s.addText(m.desc, {x:1.72, y:y+0.48, w:7.65, h:0.45, fontSize:13, color:WH, align:"right", rtlMode:true});
  });
}

// ==================================================
// SLIDE 20 - APPROACH OF THE SALAF
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: CR};
  frame(s);
  topBar(s, "منهجُ السَّلفِ في الأسماءِ والصِّفاتِ");
  s.addText("كيف تعامل النبيُّ ﷺ وصحابتُه مع نصوص الأسماء والصفات؟", {
    x:0.5, y:1.1, w:9, h:0.5, fontSize:16, bold:true, color:GM, align:"right", rtlMode:true
  });
  div(s, 1.68, 1.5, 7);
  
  const principles = [
    {icon:"✅", title:"التسليم والتعظيم", desc:"آمنوا بما جاء في القرآن والسنة دون تكلف أو تحريف"},
    {icon:"✅", title:"الفهم باللغة العربية", desc:"فهموا النصوص بمقتضى اللغة العربية الواضحة"},
    {icon:"✅", title:"نفي المشابهة", desc:"أثبتوا الصفات لله مع القطع بأنها لا تشبه صفات المخلوقين"},
    {icon:"❌", title:"لم يتكلفوا التأويل", desc:"لم يقولوا إن نصوص الصفات من المتشابه الذي لا يُفهَم"},
  ];
  
  principles.forEach((p, i) => {
    const y = 1.88 + i * 0.88;
    const bg = p.icon === "✅" ? "EEF6F1" : "FFF0F0";
    const border = p.icon === "✅" ? GM : "CC3333";
    s.addShape(pres.shapes.RECTANGLE, {x:0.45, y, w:9.1, h:0.75, fill:{color:bg}, line:{color:border, width:1}});
    s.addText(p.icon, {x:0.5, y, w:0.6, h:0.75, fontSize:20, align:"center", valign:"middle"});
    s.addText(p.title, {x:1.15, y, w:2.8, h:0.75, fontSize:15, bold:true, color:TX, align:"right", valign:"middle", rtlMode:true});
    s.addText(p.desc, {x:4.05, y, w:5.4, h:0.75, fontSize:14, color:"444444", align:"right", valign:"middle", rtlMode:true});
  });
  
  s.addShape(pres.shapes.RECTANGLE, {x:0.45, y:5.05, w:9.1, h:0.4, fill:{color:GD}, line:{width:0}});
  s.addText("﴿ لَيْسَ كَمِثْلِهِ شَيْءٌ وَهُوَ السَّمِيعُ الْبَصِيرُ ﴾ [الشورى: ١١]", {
    x:0.45, y:5.05, w:9.1, h:0.4, fontSize:15, bold:true, color:GO, align:"center", valign:"middle", rtlMode:true, italic:true
  });
}

// ==================================================
// SLIDE 21 - EXAMPLE: LOVE ATTRIBUTE
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: CR};
  frame(s);
  topBar(s, "مثالٌ تطبيقي: صفةُ المَحَبَّة");
  s.addText("﴿ إِنَّ اللَّهَ يُحِبُّ التَّوَّابِينَ ﴾ [البقرة: ٢٢٢]", {
    x:0.5, y:1.1, w:9, h:0.52, fontSize:20, bold:true, color:GM, italic:true, align:"center", rtlMode:true
  });
  div(s, 1.7, 1.5, 7);
  
  // Two columns: correct vs wrong
  s.addShape(pres.shapes.RECTANGLE, {x:0.4, y:1.85, w:4.45, h:0.5, fill:{color:GD}, line:{width:0}});
  s.addText("✅ منهج السلف", {x:0.4, y:1.85, w:4.45, h:0.5, fontSize:18, bold:true, color:GO, align:"center", valign:"middle"});
  s.addShape(pres.shapes.RECTANGLE, {x:5.15, y:1.85, w:4.45, h:0.5, fill:{color:"8B0000"}, line:{width:0}});
  s.addText("❌ منهج المُؤوِّلين", {x:5.15, y:1.85, w:4.45, h:0.5, fontSize:18, bold:true, color:WH, align:"center", valign:"middle"});
  
  s.addShape(pres.shapes.RECTANGLE, {x:0.4, y:2.38, w:4.45, h:2.55, fill:{color:"EEF6F1"}, line:{color:GM, width:1.5}});
  const correctPoints = ["نُثبِت المحبةَ صفةً حقيقية لله","محبةٌ تليق بجلاله وعظمته","لا نقول فيها غير ما جاء في اللغة العربية","﴿ لَيْسَ كَمِثْلِهِ شَيْءٌ ﴾ فمحبتُه لا تشبه محبة المخلوقين"];
  correctPoints.forEach((p, i) => {
    s.addShape(pres.shapes.RECTANGLE, {x:4.45, y:2.5+i*0.58, w:0.07, h:0.07, rotate:45, fill:{color:GM}, line:{width:0}});
    s.addText(p, {x:0.5, y:2.44+i*0.58, w:3.85, h:0.5, fontSize:13, color:TX, align:"right", valign:"middle", rtlMode:true});
  });
  
  s.addShape(pres.shapes.RECTANGLE, {x:5.15, y:2.38, w:4.45, h:2.55, fill:{color:"FFF5F5"}, line:{color:"CC3333", width:1.5}});
  const wrongPoints = ["قالوا: المحبةُ ميلُ القلب، وهذا لا يليق بالله","فأوَّلوها وقالوا: المرادُ بها الإرادة أو الثواب","في نفس الوقت أثبتوا صفة الإرادة!","لماذا؟ التناقضُ واضح!"];
  wrongPoints.forEach((p, i) => {
    s.addShape(pres.shapes.RECTANGLE, {x:5.2, y:2.5+i*0.58, w:0.07, h:0.07, rotate:45, fill:{color:"CC3333"}, line:{width:0}});
    s.addText(p, {x:5.28, y:2.44+i*0.58, w:4.2, h:0.5, fontSize:13, color:TX, align:"right", valign:"middle", rtlMode:true});
  });
  
  s.addShape(pres.shapes.RECTANGLE, {x:0.4, y:5.02, w:9.2, h:0.42, fill:{color:GM}, line:{width:0}});
  s.addText("نُثبِت الصفاتِ لله كما جاءت، مع نفي المشابهة التامَّة لصفات المخلوقين", {
    x:0.4, y:5.02, w:9.2, h:0.42, fontSize:14, bold:true, color:WH, align:"center", valign:"middle", rtlMode:true
  });
}

// ==================================================
// SLIDE 22 - ACTIVITY: BEAUTIFUL NAMES
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: GM};
  frame(s);
  s.addShape(pres.shapes.RECTANGLE, {x:0, y:0, w:10, h:1.0, fill:{color:GD}, line:{width:0}});
  s.addShape(pres.shapes.RECTANGLE, {x:0, y:0.97, w:10, h:0.05, fill:{color:GO}, line:{width:0}});
  s.addText("🌸  نشاطٌ جماعي  🌸", {x:0.5, y:0.1, w:9, h:0.78, fontSize:28, bold:true, color:GO, align:"center"});
  
  s.addText("تأمَّل هذه الأسماءَ الحُسنى وصِل كلَّ اسمٍ بمعناه!", {
    x:0.5, y:1.1, w:9, h:0.55, fontSize:22, bold:true, color:WH, align:"center", rtlMode:true
  });
  div(s, 1.75, 1.5, 7);
  
  const names = [
    ["الرَّحمن", "الرَّحيم", "الكريم", "العليم", "القادر", "الحكيم"],
    ["واسعُ الرحمة", "رحيمٌ بعباده", "كثيرُ العطاء", "يعلم كل شيء", "لا يعجزه شيء", "حكيمٌ في أفعاله"]
  ];
  
  s.addText("الأسماء:", {x:0.5, y:1.92, w:9, h:0.38, fontSize:15, bold:true, color:GX, align:"right", rtlMode:true});
  names[0].forEach((name, i) => {
    crd(s, 0.45+i*1.58, 2.35, 1.45, 0.68, GD, name, GO, 17, GO, true);
  });
  
  s.addText("المعاني:", {x:0.5, y:3.15, w:9, h:0.38, fontSize:15, bold:true, color:GX, align:"right", rtlMode:true});
  names[1].forEach((meaning, i) => {
    crd(s, 0.45+i*1.58, 3.58, 1.45, 0.85, GM, meaning, WH, 12, GO);
  });
  
  s.addText("⏱ ناقش مع زميلك ثم أجب!", {x:0.5, y:4.65, w:9, h:0.42, fontSize:16, color:GX, align:"center", rtlMode:true});
}

// ==================================================
// SLIDE 23 - SECTION HEADER: ULUHIYYAH
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: GD};
  frame(s);
  s.addText("❋", {x:4.5, y:0.8, w:1, h:0.7, fontSize:36, color:GO, align:"center"});
  s.addText("المحورُ الثالث", {x:0.5, y:1.52, w:9, h:0.58, fontSize:22, color:GX, align:"center", rtlMode:true});
  s.addText("الأُلُوهِيَّة", {
    x:0.5, y:2.1, w:9, h:1.0, fontSize:54, bold:true, color:GO, align:"center", rtlMode:true, fontFace:"Arial"
  });
  div(s, 3.2, 2, 6);
  s.addText("إفرادُ اللهِ تعالى بالتعبُّد", {
    x:0.5, y:3.35, w:9, h:0.55, fontSize:22, color:WH, align:"center", rtlMode:true
  });
  s.addText("لا إِلَهَ إِلَّا اللهُ", {
    x:0.5, y:3.95, w:9, h:0.65, fontSize:28, bold:true, color:GO, align:"center", rtlMode:true
  });
  s.addText("✦  ✦  ✦", {x:0.5, y:4.78, w:9, h:0.38, fontSize:18, color:GO, align:"center"});
}

// ==================================================
// SLIDE 24 - MEANING OF ULUHIYYAH
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: CR};
  frame(s);
  topBar(s, "معنى توحيدِ الأُلوهيةِ");
  s.addText("إفرادُ اللهِ تعالى بالعبادةِ بأَلَّا يُعبَد إلا هو", {
    x:0.5, y:1.1, w:9, h:0.5, fontSize:18, bold:true, color:GM, align:"center", rtlMode:true
  });
  div(s, 1.68, 1.5, 7);
  
  const acts = [
    {icon:"🙏", text:"لا يُسجَد إلا له"},
    {icon:"🤲", text:"لا يُدعى إلا هو"},
    {icon:"💙", text:"لا يُستغاث إلا به"},
    {icon:"🌟", text:"لا يُتوكَّل إلا عليه"},
    {icon:"❤️", text:"لا يُحَبُّ كحبِّه أحد"},
    {icon:"😰", text:"لا يُخاف خوفَ التعظيم إلا منه"},
  ];
  acts.forEach((a, i) => {
    const col = i % 2, row = Math.floor(i/2);
    const x = col === 0 ? 5.05 : 0.45;
    const y = 1.88 + row * 1.14;
    s.addShape(pres.shapes.RECTANGLE, {x, y, w:4.4, h:0.98, fill:{color:GM}, line:{color:GO, width:1.5}, shadow:{type:"outer", blur:4, offset:2, angle:135, color:"000000", opacity:0.15}});
    s.addText(a.icon, {x, y, w:0.75, h:0.98, fontSize:24, align:"center", valign:"middle"});
    s.addText(a.text, {x:x+0.75, y, w:3.55, h:0.98, fontSize:16, bold:true, color:WH, align:"right", valign:"middle", rtlMode:true});
  });
  
  s.addShape(pres.shapes.RECTANGLE, {x:0.4, y:5.08, w:9.2, h:0.38, fill:{color:GO}, line:{width:0}});
  s.addText("﴿ إِنَّ رَبَّكُمُ اللَّهُ فَاعْبُدُوهُ ﴾ [آل عمران: ٥١]", {
    x:0.4, y:5.08, w:9.2, h:0.38, fontSize:14, bold:true, color:GD, align:"center", valign:"middle", rtlMode:true, italic:true
  });
}

// ==================================================
// SLIDE 25 - RUBUBIYYAH vs ULUHIYYAH
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: CR};
  frame(s);
  topBar(s, "الفرقُ بين توحيدِ الربوبيةِ وتوحيدِ الألوهيةِ");
  
  // Left: Rububiyyah
  s.addShape(pres.shapes.RECTANGLE, {x:0.4, y:1.08, w:4.45, h:0.62, fill:{color:GM}, line:{width:0}});
  s.addText("توحيدُ الرُّبوبيةِ", {x:0.4, y:1.08, w:4.45, h:0.62, fontSize:22, bold:true, color:GO, align:"center", valign:"middle"});
  s.addShape(pres.shapes.RECTANGLE, {x:0.4, y:1.72, w:4.45, h:3.5, fill:{color:"EEF6F1"}, line:{color:GM, width:1.5}});
  const rubPoints = ["الإقرارُ بأن الله هو الخالقُ الوحيد","وأنه المالكُ المُتصرِّف في الكون","وأنه الرازقُ لجميع المخلوقات","حتى المشركون كانوا مُقِرِّين بهذا! ﴿وَلَئِن سَأَلْتَهُم مَّنْ خَلَقَهُمْ...﴾"];
  rubPoints.forEach((p, i) => {
    s.addShape(pres.shapes.RECTANGLE, {x:4.42, y:1.88+i*0.8, w:0.08, h:0.08, rotate:45, fill:{color:GM}, line:{width:0}});
    s.addText(p, {x:0.55, y:1.82+i*0.8, w:3.75, h:0.72, fontSize:13, color:TX, align:"right", valign:"middle", rtlMode:true});
  });
  
  // VS
  s.addShape(pres.shapes.OVAL, {x:4.53, y:2.65, w:0.94, h:0.94, fill:{color:GO}, line:{width:0}});
  s.addText("+", {x:4.53, y:2.65, w:0.94, h:0.94, fontSize:26, bold:true, color:GD, align:"center", valign:"middle"});
  
  // Right: Uluhiyyah
  s.addShape(pres.shapes.RECTANGLE, {x:5.15, y:1.08, w:4.45, h:0.62, fill:{color:GD}, line:{width:0}});
  s.addText("توحيدُ الأُلوهيةِ", {x:5.15, y:1.08, w:4.45, h:0.62, fontSize:22, bold:true, color:GO, align:"center", valign:"middle"});
  s.addShape(pres.shapes.RECTANGLE, {x:5.15, y:1.72, w:4.45, h:3.5, fill:{color:"F0F5FF"}, line:{color:GD, width:1.5}});
  const uluPoints = ["إفرادُ الله بالعبادة وحده","عدم توجيه أي عبادة لغيره","هو المقصودُ من بعثة الرسل","هو معنى كلمة «لا إله إلا الله»"];
  uluPoints.forEach((p, i) => {
    s.addShape(pres.shapes.RECTANGLE, {x:5.2, y:1.88+i*0.8, w:0.08, h:0.08, rotate:45, fill:{color:GD}, line:{width:0}});
    s.addText(p, {x:5.28, y:1.82+i*0.8, w:4.2, h:0.72, fontSize:13, color:TX, align:"right", valign:"middle", rtlMode:true});
  });
  
  s.addShape(pres.shapes.RECTANGLE, {x:0.4, y:5.08, w:9.2, h:0.38, fill:{color:GD}, line:{width:0}});
  s.addText("الربوبيةُ تستلزمُ الألوهيةَ: إذا آمنتَ أنه الخالقُ الوحيد → وجب أن يُعبَد وحده", {
    x:0.4, y:5.08, w:9.2, h:0.38, fontSize:13, bold:true, color:GO, align:"center", valign:"middle", rtlMode:true
  });
}

// ==================================================
// SLIDE 26 - MEANING OF LA ILAHA ILLA ALLAH
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: GD};
  frame(s);
  ornament(s, 4.6, 0.28, 24);
  s.addText("«لا إِلَهَ إِلَّا اللهُ»", {
    x:0.5, y:0.65, w:9, h:0.75, fontSize:36, bold:true, color:GO, align:"center", rtlMode:true
  });
  div(s, 1.48, 1.5, 7);
  
  s.addShape(pres.shapes.RECTANGLE, {x:0.5, y:1.62, w:9, h:1.32, fill:{color:GM}, line:{color:GO, width:2}});
  s.addText("«إله» في أصل اللغة تعني: ما يَشتاق إليه القلبُ ويتعلَّق به بمحبةٍ وتذلُّل", {
    x:0.6, y:1.68, w:8.8, h:1.12, fontSize:19, color:GX, align:"center", valign:"middle", rtlMode:true
  });
  
  div(s, 3.05, 1.5, 7);
  
  s.addText("فقول المؤمن «لا إله إلا الله» يعني:", {
    x:0.5, y:3.18, w:9, h:0.48, fontSize:17, bold:true, color:GX, align:"right", rtlMode:true
  });
  
  const meanings2 = [
    "لا محبوبَ حقيقيٌّ إلا الله",
    "لا مرهوبَ يستحق الخوفَ والتعظيمَ إلا الله",
    "لا شيءَ يملأ القلبَ ويُشبِع الروحَ إلا الله",
  ];
  meanings2.forEach((m, i) => {
    s.addShape(pres.shapes.RECTANGLE, {x:0.5, y:3.75+i*0.55, w:9, h:0.47, fill:{color:i%2===0?GM:"3D7A5A"}, line:{width:0}});
    s.addShape(pres.shapes.RECTANGLE, {x:9.08, y:3.77+i*0.55, w:0.08, h:0.08, rotate:45, fill:{color:GO}, line:{width:0}});
    s.addText(m, {x:0.55, y:3.75+i*0.55, w:8.4, h:0.47, fontSize:16, color:WH, align:"right", valign:"middle", rtlMode:true});
  });
}

// ==================================================
// SLIDE 27 - TRUE MEANING OF WORSHIP
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: CR};
  frame(s);
  topBar(s, "حقيقةُ التعبُّدِ في الإسلام");
  s.addText("العبادةُ في الإسلام ليست فقط ركوعاً وسجوداً وصياماً...", {
    x:0.5, y:1.1, w:9, h:0.5, fontSize:17, bold:true, color:GM, align:"right", rtlMode:true, italic:true
  });
  div(s, 1.67, 1.5, 7);
  s.addText("بل حقيقةُ العبادةِ هي:", {x:0.5, y:1.78, w:9, h:0.45, fontSize:17, bold:true, color:TX, align:"right", rtlMode:true});
  
  s.addShape(pres.shapes.RECTANGLE, {x:0.45, y:2.3, w:9.1, h:1.05, fill:{color:GD}, line:{color:GO, width:2}});
  s.addText("💙  محبةٌ وتذلُّلٌ وانقيادٌ للهِ تعالى", {
    x:0.55, y:2.3, w:8.9, h:1.05, fontSize:26, bold:true, color:GO, align:"center", valign:"middle", rtlMode:true
  });
  
  div(s, 3.45, 1.5, 7);
  s.addText("والأعمالُ الظاهرة (الصلاة والزكاة والصيام) لا تُقبَل ولا تنفع إذا لم يصاحبها إخلاصٌ قلبيٌّ لله؛ لأن:", {
    x:0.5, y:3.57, w:9, h:0.65, fontSize:14, color:TX, align:"right", rtlMode:true
  });
  
  s.addShape(pres.shapes.RECTANGLE, {x:0.45, y:4.28, w:9.1, h:1.12, fill:{color:GM}, line:{color:GO, width:1.5}});
  s.addText("أصلَ العبادةِ الذي تنشأ عنه الأعمالُ الظاهرةُ هو: الاعترافُ القلبيُّ والانقيادُ الباطنيُّ والتسليمُ لله تعالى", {
    x:0.6, y:4.28, w:8.9, h:1.12, fontSize:17, color:WH, align:"center", valign:"middle", rtlMode:true
  });
}

// ==================================================
// SLIDE 28 - LOVE AS WORSHIP
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: GD};
  frame(s);
  s.addText("مِن أعظمِ صُوَرِ العبادةِ: المَحَبَّة", {
    x:0.5, y:0.62, w:9, h:0.65, fontSize:28, bold:true, color:GO, align:"center", rtlMode:true
  });
  div(s, 1.35, 1.5, 7);
  
  const loveTypes = [
    {icon:"✨", text:"حبُّه لكماله وجلاله وجماله"},
    {icon:"👑", text:"حبُّه لربوبيته وعظمته"},
    {icon:"🎁", text:"حبُّه لإنعامه علينا"},
    {icon:"📖", text:"حبُّه لإرساله الرسلَ وبيانِ الغاية من وجودنا"},
  ];
  loveTypes.forEach((l, i) => {
    s.addShape(pres.shapes.RECTANGLE, {x:0.5, y:1.52+i*0.92, w:9, h:0.78, fill:{color:GM}, line:{color:GO, width:1}});
    s.addText(l.icon, {x:0.55, y:1.52+i*0.92, w:0.75, h:0.78, fontSize:24, align:"center", valign:"middle"});
    s.addText(l.text, {x:1.38, y:1.52+i*0.92, w:8.0, h:0.78, fontSize:17, color:WH, align:"right", valign:"middle", rtlMode:true});
  });
  
  s.addShape(pres.shapes.RECTANGLE, {x:0.5, y:5.1, w:9, h:0.38, fill:{color:GO}, line:{width:0}});
  s.addText("كلّما ازداد العبدُ محبةً لربِّه → طاب قلبُه به، واستغنى به عن غيره، وذاق لذَّةَ الإيمان", {
    x:0.5, y:5.1, w:9, h:0.38, fontSize:13, bold:true, color:GD, align:"center", valign:"middle", rtlMode:true
  });
}

// ==================================================
// SLIDE 29 - FRUITS OF TAWHEED
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: CR};
  frame(s);
  topBar(s, "ثمراتُ التوحيدِ في الحياةِ");
  
  const fruits = [
    {icon:"🌱", title:"التحرُّر من الخوف", desc:"الموحِّدُ لا يخشى إلا الله، فيعيش مُطمئنَّاً حراً"},
    {icon:"💎", title:"الكرامةُ والعزَّة", desc:"لا يتذلل لغير الله، فيعيش كريماً عزيزاً"},
    {icon:"🧭", title:"وضوحُ الهدف", desc:"يعرف لماذا خُلِق وإلى أين يسير"},
    {icon:"🤲", title:"دوامُ التوكُّل", desc:"يتوكَّل على الله في كل أمر فتطمئن نفسه"},
    {icon:"⚖️", title:"توازنُ الشخصية", desc:"لا إفراطٌ في التعلق بالدنيا ولا قنوطٌ من رحمة الله"},
    {icon:"🌟", title:"الاستعلاء بالإيمان", desc:"يشعر بالأمان والقوة في مواجهة المواقف الصعبة"},
  ];
  
  fruits.forEach((f, i) => {
    const col = i % 2, row = Math.floor(i/2);
    const x = col === 0 ? 5.05 : 0.45;
    const y = 1.08 + row * 1.45;
    s.addShape(pres.shapes.RECTANGLE, {x, y, w:4.4, h:1.28, fill:{color:GM}, line:{color:GO, width:1.5}, shadow:{type:"outer", blur:4, offset:2, angle:135, color:"000000", opacity:0.15}});
    s.addText(f.icon, {x, y, w:0.8, h:1.28, fontSize:28, align:"center", valign:"middle"});
    s.addText(f.title, {x:x+0.85, y:y+0.05, w:3.45, h:0.45, fontSize:16, bold:true, color:GX, align:"right", rtlMode:true});
    s.addText(f.desc, {x:x+0.85, y:y+0.52, w:3.45, h:0.7, fontSize:13, color:WH, align:"right", rtlMode:true});
  });
}

// ==================================================
// SLIDE 30 - SIGNS OF WEAK FAITH: SUPERSTITIONS
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: CR};
  frame(s);
  topBar(s, "من علاماتِ ضَعفِ العقيدة: التعلُّق بالخُرافات");
  s.addText("إذا آمن الإنسانُ بالتوحيد الحقيقي → لا يتعلق بغير الله في جلب النفع أو دفع الضر", {
    x:0.5, y:1.1, w:9, h:0.55, fontSize:15, color:TX, align:"right", rtlMode:true
  });
  div(s, 1.72, 1.5, 7);
  
  // Examples of superstitions
  s.addShape(pres.shapes.RECTANGLE, {x:0.45, y:1.87, w:9.1, h:0.55, fill:{color:"CC3333"}, line:{width:0}});
  s.addText("❌  أمثلةٌ على التعلق بالخرافات (من أنواع الشرك)", {
    x:0.45, y:1.87, w:9.1, h:0.55, fontSize:16, bold:true, color:WH, align:"center", valign:"middle", rtlMode:true
  });
  
  const superstitions = [
    {icon:"👁️", text:"تعليقُ صورة العين لدفع الحسد"},
    {icon:"📿", text:"ارتداءُ التمائم والقلائد للحماية والوقاية"},
    {icon:"🙏", text:"الاستغاثةُ بالأموات: «يا حسين! يا عبد القادر!»"},
    {icon:"🧿", text:"تعليقُ أشياء في السيارة أو البيت لدفع الشر"},
  ];
  superstitions.forEach((sup, i) => {
    const y = 2.52 + i * 0.66;
    s.addShape(pres.shapes.RECTANGLE, {x:0.45, y, w:9.1, h:0.58, fill:{color:"FFF5F5"}, line:{color:"CC3333", width:0.8}});
    s.addText(sup.icon, {x:0.5, y, w:0.65, h:0.58, fontSize:22, align:"center", valign:"middle"});
    s.addText(sup.text, {x:1.22, y, w:8.2, h:0.58, fontSize:15, color:TX, align:"right", valign:"middle", rtlMode:true});
  });
  
  s.addShape(pres.shapes.RECTANGLE, {x:0.45, y:5.12, w:9.1, h:0.38, fill:{color:GD}, line:{width:0}});
  s.addText("قال النبي ﷺ: «مَن عَلَّق تميمةً فقد أَشرَك» [أحمد]", {
    x:0.45, y:5.12, w:9.1, h:0.38, fontSize:14, bold:true, color:GO, align:"center", valign:"middle", rtlMode:true, italic:true
  });
}

// ==================================================
// SLIDE 31 - SIGNS OF WEAK FAITH: SWEARING
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: CR};
  frame(s);
  topBar(s, "من علاماتِ ضَعفِ العقيدة: الحَلِفُ بغيرِ اللهِ");
  
  s.addShape(pres.shapes.RECTANGLE, {x:0.45, y:1.08, w:9.1, h:1.25, fill:{color:GD}, line:{color:GO, width:2}});
  s.addText("قال النبي ﷺ: «أَلَا مَن كَانَ حَالِفاً فَلَا يَحلِف إِلَّا بِاللهِ»", {
    x:0.55, y:1.08, w:8.9, h:1.25, fontSize:22, bold:true, color:GO, italic:true, align:"center", valign:"middle", rtlMode:true
  });
  s.addText("متفق عليه", {x:0.5, y:2.42, w:9, h:0.35, fontSize:14, color:GL, align:"center", rtlMode:true});
  
  div(s, 2.85, 1.5, 7);
  
  s.addShape(pres.shapes.RECTANGLE, {x:0.45, y:2.98, w:9.1, h:0.52, fill:{color:"CC3333"}, line:{width:0}});
  s.addText("❌  أمثلةٌ لا تجوز شرعاً:", {
    x:0.45, y:2.98, w:9.1, h:0.52, fontSize:16, bold:true, color:WH, align:"center", valign:"middle", rtlMode:true
  });
  
  const wrongOaths = [
    "الحَلِفُ بالشَّرَفِ أو العِرضِ",
    "الحَلِفُ بحياةِ فلانٍ أو حياتي",
    "الحَلِفُ بالنبيِّ أو الصالحين",
    "الحَلِفُ بالوطنِ أو العَلَم",
  ];
  wrongOaths.forEach((oath, i) => {
    s.addShape(pres.shapes.RECTANGLE, {x:0.45, y:3.58+i*0.4, w:9.1, h:0.35, fill:{color:i%2===0?"FFF5F5":"FFF0F0"}, line:{color:"CC3333", width:0.5}});
    s.addText("❌  " + oath, {x:0.55, y:3.58+i*0.4, w:8.9, h:0.35, fontSize:15, color:"990000", align:"right", valign:"middle", rtlMode:true});
  });
  
  s.addShape(pres.shapes.RECTANGLE, {x:0.45, y:5.12, w:9.1, h:0.38, fill:{color:GM}, line:{width:0}});
  s.addText("✅  الحَلِفُ بالله وحده هو المشروع: «واللهِ، وتاللهِ، وبِاللهِ»", {
    x:0.45, y:5.12, w:9.1, h:0.38, fontSize:14, bold:true, color:WH, align:"center", valign:"middle", rtlMode:true
  });
}

// ==================================================
// SLIDE 32 - GROUP DISCUSSION ACTIVITY
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: GM};
  frame(s);
  s.addShape(pres.shapes.RECTANGLE, {x:0, y:0, w:10, h:1.0, fill:{color:GD}, line:{width:0}});
  s.addShape(pres.shapes.RECTANGLE, {x:0, y:0.97, w:10, h:0.05, fill:{color:GO}, line:{width:0}});
  s.addText("🌿  نشاطٌ جماعي ختامي  🌿", {x:0.5, y:0.1, w:9, h:0.78, fontSize:26, bold:true, color:GO, align:"center"});
  
  s.addText("التوحيدُ في حياتنا اليومية", {x:0.5, y:1.12, w:9, h:0.58, fontSize:24, bold:true, color:WH, align:"center", rtlMode:true});
  div(s, 1.78, 1.5, 7);
  
  s.addText("⏱ الوقت: ٧ دقائق للنقاش ثم العرض", {x:0.5, y:1.92, w:9, h:0.42, fontSize:16, color:GX, align:"center", rtlMode:true});
  
  const questions = [
    {q:"السؤالُ الأول:", text:"كيف يظهر التوحيدُ في دعائك ليلاً عند الضيق؟"},
    {q:"السؤالُ الثاني:", text:"ما الفرق بين من يدعو اللهَ مباشرةً ومن يذهب إلى الأضرحة؟"},
    {q:"السؤالُ الثالث:", text:"ما الذي يجب أن تغيره في حياتك اليومية بعد فهم التوحيد؟"},
  ];
  questions.forEach((q, i) => {
    s.addShape(pres.shapes.RECTANGLE, {x:0.45, y:2.45+i*1.0, w:9.1, h:0.88, fill:{color:GD}, line:{color:GO, width:1.5}});
    s.addText(q.q, {x:0.55, y:2.5+i*1.0, w:8.9, h:0.32, fontSize:14, bold:true, color:GO, align:"right", rtlMode:true});
    s.addText(q.text, {x:0.55, y:2.82+i*1.0, w:8.9, h:0.45, fontSize:15, color:WH, align:"right", valign:"middle", rtlMode:true});
  });
}

// ==================================================
// SLIDE 33 - OVERALL SUMMARY
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: GD};
  frame(s);
  ornament(s, 4.6, 0.28, 24);
  s.addText("خُلاصةُ المحاضرةِ", {
    x:0.5, y:0.65, w:9, h:0.65, fontSize:32, bold:true, color:GO, align:"center", rtlMode:true
  });
  div(s, 1.38, 1.5, 7);
  
  const summary = [
    {num:"١", title:"وجودُ الله", points:"يثبت بالعقل (الاستدلال بالأثر) وبالفطرة (المعارف الأولية - الافتقار - الغرائز)"},
    {num:"٢", title:"أسماءُ الله الحسنى", points:"اعرفها واحفظها وتعبَّد لله بها وادعُه بها، مع إثبات الصفات دون تشبيه"},
    {num:"٣", title:"توحيدُ الألوهية", points:"أفرِد الله بالعبادة في كل شيء، ولا تتعلق بمخلوق ولا بخرافة"},
  ];
  summary.forEach((item, i) => {
    s.addShape(pres.shapes.RECTANGLE, {x:0.45, y:1.55+i*1.28, w:9.1, h:1.12, fill:{color:GM}, line:{color:GO, width:1.5}});
    s.addShape(pres.shapes.RECTANGLE, {x:0.45, y:1.55+i*1.28, w:0.85, h:1.12, fill:{color:GO}, line:{width:0}});
    s.addText(item.num, {x:0.45, y:1.55+i*1.28, w:0.85, h:1.12, fontSize:24, bold:true, color:GD, align:"center", valign:"middle"});
    s.addText(item.title, {x:1.38, y:1.6+i*1.28, w:7.95, h:0.38, fontSize:18, bold:true, color:GX, align:"right", rtlMode:true});
    s.addText(item.points, {x:1.38, y:1.98+i*1.28, w:7.95, h:0.65, fontSize:13, color:WH, align:"right", rtlMode:true});
  });
  
  s.addText("✦  ✦  ✦", {x:0.5, y:5.28, w:9, h:0.25, fontSize:14, color:GO, align:"center"});
}

// ==================================================
// SLIDE 34 - REVIEW QUESTIONS
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: CR};
  frame(s);
  topBar(s, "مراجعةٌ سريعة 🎯");
  
  const questions = [
    {q:"كم عددُ أركان الإيمان؟ واذكرها.", level:"سهل"},
    {q:"ما معنى «الاستدلال بالأثر على المؤثر»؟ اضرب مثالاً.", level:"متوسط"},
    {q:"ما الفرق بين توحيد الربوبية وتوحيد الألوهية؟", level:"متوسط"},
    {q:"اذكر ثلاثةَ جهات تدلُّ فيها الفطرةُ على وجود الله.", level:"متقدم"},
    {q:"من علامات ضعف العقيدة: اذكر مثالين.", level:"سهل"},
  ];
  
  const levelColors = {سهل: "4A9B6F", متوسط: "C9A227", متقدم: "8B0000"};
  const levelBg = {سهل: "EEF6F1", متوسط: "FFF9E6", متقدم: "FFF0F0"};
  
  questions.forEach((q, i) => {
    const y = 1.1 + i * 0.87;
    s.addShape(pres.shapes.RECTANGLE, {x:0.45, y, w:9.1, h:0.75, fill:{color:levelBg[q.level]}, line:{color:levelColors[q.level], width:1}});
    s.addShape(pres.shapes.RECTANGLE, {x:0.45, y, w:1.2, h:0.75, fill:{color:levelColors[q.level]}, line:{width:0}});
    s.addText(q.level, {x:0.45, y, w:1.2, h:0.75, fontSize:13, bold:true, color:WH, align:"center", valign:"middle", rtlMode:true});
    s.addText(`${i+1}. ${q.q}`, {x:1.72, y, w:7.65, h:0.75, fontSize:14, color:TX, align:"right", valign:"middle", rtlMode:true});
  });
}

// ==================================================
// SLIDE 35 - TIMELINE / LESSON PLAN
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: CR};
  frame(s);
  topBar(s, "توزيعُ الوقتِ — محاضرةُ ساعةٍ واحدة");
  
  const timing = [
    {time:"٥ د", section:"مقدمة", detail:"أركان الإيمان الستة وأهمية الإيمان بالله", color: GD},
    {time:"٢٠ د", section:"المحور الأول", detail:"وجود الله: دليل العقل والفطرة + نشاط جماعي", color: GM},
    {time:"١٥ د", section:"المحور الثاني", detail:"أسماء الله وصفاته + نشاط الأسماء الحسنى", color: GO},
    {time:"١٥ د", section:"المحور الثالث", detail:"الألوهية + ثمرات التوحيد + علامات الضعف", color: GM},
    {time:"٥ د", section:"الخاتمة", detail:"مراجعة + أسئلة + نشاط ختامي", color: GD},
  ];
  
  timing.forEach((t, i) => {
    const y = 1.1 + i * 0.88;
    s.addShape(pres.shapes.RECTANGLE, {x:0.4, y, w:9.2, h:0.78, fill:{color:"EEF6F1"}, line:{color:GM, width:0.8}});
    s.addShape(pres.shapes.RECTANGLE, {x:0.4, y, w:1.05, h:0.78, fill:{color:t.color}, line:{width:0}});
    s.addText(t.time, {x:0.4, y, w:1.05, h:0.78, fontSize:16, bold:true, color: t.color === GO ? GD : GO, align:"center", valign:"middle", rtlMode:true});
    s.addText(t.section, {x:1.52, y, w:2.5, h:0.78, fontSize:15, bold:true, color:TX, align:"right", valign:"middle", rtlMode:true});
    s.addText(t.detail, {x:4.1, y, w:5.4, h:0.78, fontSize:14, color:"444444", align:"right", valign:"middle", rtlMode:true});
  });
  
  s.addShape(pres.shapes.RECTANGLE, {x:0.4, y:5.5, w:9.2, h:0.0, fill:{color:GO}, line:{width:0}});
}

// ==================================================
// SLIDE 36 - REFERENCES
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: CR};
  frame(s);
  topBar(s, "مراجعُ للاستزادة");
  
  const refs = [
    {icon:"📗", title:"البناء العقدي للجيل الصاعد", author:"الكتابُ الأصلي للمحاضرة"},
    {icon:"📘", title:"لأنك الله", author:"علي الفيفي — مناسبٌ جداً للشباب في أسماء الله"},
    {icon:"📙", title:"براهين وجود الله", author:"د. سامي عامري"},
    {icon:"📕", title:"الصنع المتقن", author:"مصطفى قديح"},
    {icon:"📓", title:"شموع النهار", author:"عبدالله العجيري"},
    {icon:"🌐", title:"قناة بيان العلمية وقناة سلسبيل", author:"للاستزادة في تعلم العقيدة"},
  ];
  
  refs.forEach((r, i) => {
    const col = i % 2, row = Math.floor(i/2);
    const x = col === 0 ? 5.1 : 0.45;
    const y = 1.1 + row * 1.42;
    s.addShape(pres.shapes.RECTANGLE, {x, y, w:4.42, h:1.22, fill:{color:GM}, line:{color:GO, width:1.5}});
    s.addText(r.icon, {x, y, w:0.7, h:1.22, fontSize:28, align:"center", valign:"middle"});
    s.addText(r.title, {x:x+0.75, y:y+0.05, w:3.55, h:0.52, fontSize:14, bold:true, color:GX, align:"right", valign:"middle", rtlMode:true});
    s.addText(r.author, {x:x+0.75, y:y+0.58, w:3.55, h:0.58, fontSize:12, color:WH, align:"right", valign:"middle", rtlMode:true});
  });
}

// ==================================================
// SLIDE 37 - CLOSING PRAYER
// ==================================================
{
  const s = pres.addSlide();
  s.background = {color: GD};
  frame(s);
  s.addText("✦  ❋  ✦", {x:0.5, y:0.45, w:9, h:0.55, fontSize:26, color:GO, align:"center"});
  
  s.addText("ختامُ الدرس", {x:0.5, y:1.0, w:9, h:0.62, fontSize:32, bold:true, color:GO, align:"center", rtlMode:true});
  div(s, 1.72, 2, 6);
  
  s.addShape(pres.shapes.RECTANGLE, {x:0.5, y:1.85, w:9, h:1.55, fill:{color:GM}, line:{color:GO, width:2}});
  s.addText("«اللَّهُمَّ ارزُقنا تَعظِيمَكَ وَخَشيَتَكَ، وَحُسنَ العَمَلِ بِالعِلمِ يَا رَبَّ العَالَمِينَ... آمين»", {
    x:0.65, y:1.85, w:8.7, h:1.55, fontSize:22, color:GX, italic:true, align:"center", valign:"middle", rtlMode:true
  });
  
  div(s, 3.5, 2, 6);
  s.addText("تذكَّروا:", {x:0.5, y:3.65, w:9, h:0.42, fontSize:17, bold:true, color:GX, align:"center", rtlMode:true});
  s.addText("العقيدةُ ليست مجرَّد معلوماتٍ تُحفَظ، بل هي حياةٌ يعيشها القلبُ ويترجمها السلوك", {
    x:0.5, y:4.12, w:9, h:0.72, fontSize:18, color:WH, align:"center", rtlMode:true
  });
  
  s.addText("✦  ✦  ✦", {x:0.5, y:5.1, w:9, h:0.38, fontSize:18, color:GO, align:"center"});
}

// ==================================================
// OUTPUT
// ==================================================
pres.writeFile({fileName: "/home/claude/ارکان_الایمان_عرض_تقدیمی.pptx"})
  .then(() => console.log("✅ DONE: Presentation created successfully!"))
  .catch(e => console.error("❌ Error:", e));
