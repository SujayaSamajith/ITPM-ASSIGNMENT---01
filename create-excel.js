const xlsx = require('xlsx');

// 50 NEGATIVE test cases - IT23361522
// 24 types per assignment PDF Appendix 1
// TC ID, Input length type, Input, Expected output, Actual output (blank), Status (blank)
// Singlish input types covered, Evidence or rationale

const cases = [
  // Type 1: Question forms (x3)
  ['Neg_0001','S','ada school giyanawa da?','අද school ගියනවා ද?','','','Question forms','The overall input is a question asking whether someone went to school today.'],
  ['Neg_0002','M','mama danne na oya koheda giye kiyala. api heta yamuda nathnam awarudda wages karanno?','මම දන්නේ නා ඔය කොහෙද ගියේ කියලා. අපි හෙට යමුද නැත්නම් අවුරුද්ද වාගේ කරන්නෝ?','','','Question forms; Inputs with Punctuation Marks','Contains two questions; uses ? marks; "nathnam" conditional clause creates embedded question.'],
  ['Neg_0003','M','oya API eke documentation eka balala da? mama enna one nam ticket ekak gannada?','ඔය API එකේ documentation එකේ බලලා ද? මම එන්නේ ඕනේ නම් ticket එකක් ගන්නද?','','','Question forms; English Digital Terms in Singlish','Two-question input; digital terms API and documentation embedded in Singlish questions.'],

  // Type 2: Command forms (x2)
  ['Neg_0004','M','wahama gedara yanna! eppa inna.','වහාම ගෙදර යන්නේ! එපා ඉන්නේ.','','','Command forms; Inputs with Punctuation Marks','Direct command with "wahama" (immediately) + negation "eppa". Uses exclamation mark.'],
  ['Neg_0005','M','oka karannam kiyala kiyanna eppa. mama kiyanawa ehema karanna. hadilla dan enna kiyannako machan.','ඒකා කරන්නම් කියලා කියන්නේ එපා. මම කියනවා එහෙම කරන්නේ. හදිල්ලා දන් එන්නකෝ මචන්.','','','Command forms; Inputs with Slang and Casual Phrasing','Negative command + urgency command; slang "machan" and "kiyannako" casual command suffix.'],

  // Type 3: Greetings (x2)
  ['Neg_0006','S','kohomada bro? sudda da inne?','කොහොමද බ්‍රෝ? සුද්ද ද ඉන්නේ?','','','Greetings; Inputs with Slang and Casual Phrasing','Informal greeting using slang "bro". Greeting embedded with casual question about wellbeing.'],
  ['Neg_0007','M','malli! godak dawasak passe kethuwa ne. kohomada inne oya? honda wela thiyenawa wage penewa.','මල්ලි! ගොඩක් දවසක් පස්සේ කෙතුවා නේ. කොහොමද ඉන්නේ ඔය? හොඳ වෙලා තිෙනවා වාගේ පෙනෙවා.','','','Greetings; Inputs with Punctuation Marks','Extended greeting to a younger person; uses "malli" (younger brother); includes exclamation + question.'],

  // Type 4: Requests (x2)
  ['Neg_0008','M','karunakara meka poddak balanna puluwanda?','කරුණාකරා මෙකා පොඩ්ඩක් බලන්නේ පුළුවන්ද?','','','Requests','Polite request using "karunakara" (please). System often mishandles the polite particle.'],
  ['Neg_0009','M','mama oba kiyanawa, oya heta meeting ekedi Zoom link eka ewanna puluwanda? mama office eke innawa hadissiye.','මම ඔබ කියනවා, ඔය හෙට meeting එකේදී Zoom link එකේ එවන්නේ පුළුවන්ද? මම office එකේ ඉන්නවා හදිස්සියේ.','','','Requests; Platform/App Names in Singlish; Isolated English Word Insertions in Singlish','Polite request to send Zoom link; platform name Zoom; English words meeting and office embedded.'],

  // Type 5: Responses (x2)
  ['Neg_0010','M','ane hari aiye. mama eka karannam.','අනේ හරි අයිේ. මම ඒකා කරන්නම්.','','','Responses','Affirmative casual response using "ane hari aiye". System fails on "ane" exclamation mapping.'],
  ['Neg_0011','M','nathne mama danne. eth oya kiyanawa nam eka karanna. mama hithuwa karanna eth baluwe naha.','නැතනේ මම දන්නේ. ඒත් ඔය කියනවා නම් ඒකා කරන්නේ. මම හිතුව කරන්නේ ඒත් බැලුවේ නාහා.','','','Responses; Inputs with Punctuation Marks','Concessive response - denying then agreeing. Complex clause boundaries confuse system.'],

  // Type 6: Repeated Words (x2)
  ['Neg_0012','S','boru boru katha kiyanna eppa.','බොරු බොරු කතා කියන්නේ එපා.','','','Repeated Words; Command forms','Repeated word "boru boru" for emphasis + negative command. System collapses or mishandles repetition.'],
  ['Neg_0013','M','hari hari mama ennam. godak godak wage innawa kiyalane. poddak poddak karakarala mama eka hadagannam.','හරි හරි මම එන්නම්. ගොඩක් ගොඩක් වාගේ ඉන්නවා කියලානේ. පොඩ්ඩක් පොඩ්ඩක් කරාකරාලා මම ඒකා හදාගන්නම්.','','','Repeated Words','Three sets of repeated words in same input. System inconsistently handles each repetition pattern.'],

  // Type 7: Inputs with Punctuation Marks (x2)
  ['Neg_0014','S','ane! oya enawada... ne?','අනේ! ඔය එනවාද... නේ?','','','Inputs with Punctuation Marks; Responses','Ellipsis (...) combined with exclamation and question mark. System strips or misplaces ellipsis.'],
  ['Neg_0015','M','mama kiyanawa, oya danna da meka? hadanne ne... eth hari karanna one, ne! mama eka balamu.','මම කියනවා, ඔය දන්නා ද මෙකා? හදන්නේ නේ... ඒත් හරි කරන්නේ ඕනේ, නේ! මම ඒකා බලාමු.','','','Inputs with Punctuation Marks; Question forms','Multiple punctuation types (comma, question mark, ellipsis, exclamation) in single input.'],

  // Type 8: Romanization / Spelling Variants (x2)
  ['Neg_0016','S','mee ekee hondha ne machan?','මේ එකේ හොඳ නේ මචන්?','','','Romanization / Spelling Variants; Inputs with Slang and Casual Phrasing','Non-standard: "ekee" for "eke", "hondha" for "honda". System rejects non-standard spelling variants.'],
  ['Neg_0017','M','man oyta passee ennam kiyala kiwwa. eth oyt balagena inne ne? mnge gedra ahuwoth enawa nam.','මං ඔයාට පස්සේ එන්නම් කියලා කිව්වා. ඒත් ඔයාට බලාගෙන ඉන්නේ නේ? මංගෙ ගෙදර ඇහුවොත් එනවා නම්.','','','Romanization / Spelling Variants','Heavily contracted spelling: "oyt" for "oyata", "mnge" for "mange", "gedra" for "gedara". System fails on dropped vowels.'],

  // Type 9: Isolated English Word Insertions (x2)
  ['Neg_0018','S','lunch karanna yamu ne?','lunch කරන්නේ යමු නේ?','','','Isolated English Word Insertions in Singlish; Question forms','Single English word "lunch" inserted in Sinhala sentence. System may transliterate "lunch".'],
  ['Neg_0019','M','mama ada project eka finish karannam kiyala hithuwa eth deadline eka miss una nisa manager kiyanawa extension ekak ganna kiyala.','මම අද project එකේ finish කරන්නම් කියලා හිතුව ඒත් deadline එකේ miss උනා නිසා manager කියනවා extension එකක් ගන්නේ කියලා.','','','Isolated English Word Insertions in Singlish','Multiple single English words (project, finish, deadline, miss, manager, extension) embedded in Singlish text.'],

  // Type 10: Multi-Word English Phrases (x2)
  ['Neg_0020','M','mama kiyanawa out of stock wela thiyenawa eth just in case order karanna hithuwa ne.','මම කියනවා out of stock වෙලා තිෙනවා ඒත් just in case order කරන්නේ හිතුව නේ.','','','Multi-Word English Phrases in Singlish; Isolated English Word Insertions in Singlish','English phrases "out of stock" and "just in case" embedded; single word "order" also inserted.'],
  ['Neg_0021','M','oya dannawa da free of charge labhena bawa? eth terms and conditions balanna one kiyalane. as soon as possible submit karanna.','ඔය දන්නවා ද free of charge ලැබෙන බව? ඒත් terms and conditions බලන්නේ ඕනේ කියලානේ. as soon as possible submit කරන්නේ.','','','Multi-Word English Phrases in Singlish; Question forms; Inputs with Punctuation Marks','Three English multi-word phrases in same input; question form; uses ? mark.'],

  // Type 11: English Digital Terms (x2)
  ['Neg_0022','S','download link eka pennanna.','download link එකේ පෙන්නන්නේ.','','','English Digital Terms in Singlish; Requests','Digital terms "download" and "link" embedded in a request sentence.'],
  ['Neg_0023','M','oya account eka login karanna puluwan eth password reset karanna one kiyanne kochcharada? email eka confirm wewida?','ඔය account එකේ login කරන්නේ පුළුවන් ඒත් password reset කරන්නේ ඕනේ කියන්නේ කොච්චරද? email එකේ confirm වෙවිද?','','','English Digital Terms in Singlish; Question forms','Digital terms (account, login, password, reset, email, confirm) throughout; ends with question.'],

  // Type 12: Platform/App Names (x2)
  ['Neg_0024','S','Viber eken message ekak ewa.','Viber එකෙන් message එකක් එව.','','','Platform/App Names in Singlish; Requests','Platform name "Viber" + digital term "message". Request to send a message via a specific app.'],
  ['Neg_0025','M','mama Facebook eke post ekak dammata api Instagram walata share karannam kiyala kiwwa. TikTok eke trend eka ballako.','මම Facebook එකේ post එකක් දාම්ම අපි Instagram වලට share කරන්නම් කියලා කිව්වා. TikTok එකේ trend එකේ බල්ලකෝ.','','','Platform/App Names in Singlish; Isolated English Word Insertions in Singlish','Three platform names (Facebook, Instagram, TikTok) with Singlish verbs. System fails on proper noun handling.'],

  // Type 13: English Abbreviations/Acronyms (x2)
  ['Neg_0026','S','ASAP karannam ok da?','ASAP කරන්නම් ok ද?','','','English Abbreviations/Acronyms in Singlish; Question forms','Acronym "ASAP" + casual "ok". System may try to phonetically map capital letters.'],
  ['Neg_0027','M','oya CV eka update karalada? HR kiyanawa ASAP submit karanna kiyala. eth DOB saha NIC eka heri da check karanna.','ඔය CV එකේ update කරලාද? HR කියනවා ASAP submit කරන්නේ කියලා. ඒත් DOB සහ NIC එකේ හරිද check කරන්නේ.','','','English Abbreviations/Acronyms in Singlish; Question forms; Isolated English Word Insertions in Singlish','Multiple acronyms (CV, HR, ASAP, DOB, NIC) in one input; embedded questions; digital terms update/submit/check.'],

  // Type 14: English Clipped Forms (x2)
  ['Neg_0028','M','prof eken lab assignment ekak thiyenawa.','prof එකෙන් lab assignment එකක් තිෙනවා.','','','English Clipped Forms in Singlish; Isolated English Word Insertions in Singlish','"prof" (professor) and "lab" (laboratory) are clipped English forms. System may transliterate them.'],
  ['Neg_0029','M','ada uni eke lec kiyala penela. mama lib eke sit karannam kiyala hithuwa eth doc ekak print karanna went.','අද uni එකේ lec කියලා පෙනෙලා. මම lib එකේ sit කරන්නම් කියලා හිතුව ඒත් doc එකක් print කරන්නේ ගිය.','','','English Clipped Forms in Singlish; Isolated English Word Insertions in Singlish','Clipped forms: uni (university), lec (lecture), lib (library), doc (document) + single English words sit/print.'],

  // Type 15: Place Names Embedded in Singlish (x2)
  ['Neg_0030','S','Galle yamuda ada?','Galle යමුද අද?','','','Place Names Embedded in Singlish; Question forms','Sri Lankan place name "Galle" embedded in simple question. System may transliterate proper noun.'],
  ['Neg_0031','M','mama Nugegoda sita Malabe yanna one. Kesbewa pass wenakota call karannam. Boralesgamuwa hotel eke lunch karannam.','මම Nugegoda සිට Malabe යන්නේ ඕනේ. Kesbewa pass වෙනකොට call කරන්නම්. Boralesgamuwa hotel එකේ lunch කරන්නම්.','','','Place Names Embedded in Singlish; Isolated English Word Insertions in Singlish','Multiple Sri Lankan place names (Nugegoda, Malabe, Kesbewa, Boralesgamuwa); English words lunch and call embedded.'],

  // Type 16: Person Names Embedded in Singlish (x2)
  ['Neg_0032','S','Ishara heta aawada?','Ishara හෙට ආවාද?','','','Person Names Embedded in Singlish; Question forms','Person name "Ishara" embedded in a question. System may transliterate the proper name.'],
  ['Neg_0033','M','Kavindra saha Thisara dennama project eke work karala maru. eth Dinusha kiyanawa ehema ne kiyala. oya Pavithra kiyanawa eka gana?','Kavindra සහ Thisara දෙන්නාම project එකේ work කරලා මාරු. ඒත් Dinusha කියනවා එහෙම නේ කියලා. ඔය Pavithra කියනවා ඒකා ගැනා?','','','Person Names Embedded in Singlish; Isolated English Word Insertions in Singlish; Question forms','Four person names (Kavindra, Thisara, Dinusha, Pavithra) in one input; English word project/work; ends with question.'],

  // Type 17: Inputs with Numbers and Numeric Suffixes (x2)
  ['Neg_0034','M','2nd match eke score eka kiyanna.','2nd match එකේ score එකේ කියන්නේ.','','','Inputs with Numbers and Numeric Suffixes; Isolated English Word Insertions in Singlish','Ordinal suffix "2nd" + English sports terms match/score. System fails on "nd" suffix after digit.'],
  ['Neg_0035','M','science lab eke 1st group eka 98% aragena. 2nd group eka 85k wage markings. 3rd group walata 50 sita 75 dakwa.','science lab එකේ 1st group එකේ 98% අරගෙනේ. 2nd group එකේ 85k වාගේ markings. 3rd group වලට 50 සිට 75 දාකාවා.','','','Inputs with Numbers and Numeric Suffixes; English Clipped Forms in Singlish','Ordinal suffixes (1st, 2nd, 3rd) + percentage (98%) + numeric suffix "k" (85k) all in one input.'],

  // Type 18: Inputs with Currency (x2)
  ['Neg_0036','S','Rs. 1500 dennako machan.','Rs. 1500 දෙන්නකෝ මචන්.','','','Inputs with Currency; Inputs with Slang and Casual Phrasing','Currency "Rs." + amount 1500 + slang "machan". System struggles with "Rs." abbreviation and decimal point.'],
  ['Neg_0037','M','mama kiwwa LKR 4500 wage weyi kiyala. eth USD 15 kiyanne rupiyel walin kochcharada? EUR 100 kiyanne godak ne.','මම කිව්වා LKR 4500 වාගේ වේ යි කියලා. ඒත් USD 15 කියන්නේ රුපියල් වලින් කොච්චරද? EUR 100 කියන්නේ ගොඩක් නේ.','','','Inputs with Currency; Question forms','Three currency codes (LKR, USD, EUR) in single input. System treats currency abbreviations inconsistently.'],

  // Type 19: Inputs with Time Formats (x2)
  ['Neg_0038','S','4:30PM nagin enna karanna.','4:30PM නාගින් එන්නේ කරන්නේ.','','','Inputs with Time Formats; Command forms','Time "4:30PM" with colon and AM/PM suffix. Command to arrive before specific time.'],
  ['Neg_0039','M','meeting eka 9:00AM walata set karana. mama 8:45 walata enna hithuwa eth traffic nisa 9:15 walata una. tomorrow 2:00PM walata reschedule karannam.','meeting එකේ 9:00AM වලට set කරනේ. මම 8:45 වලට එන්නේ හිතුව ඒත් traffic නිසා 9:15 වලට උනා. tomorrow 2:00PM වලට reschedule කරන්නම්.','','','Inputs with Time Formats; Isolated English Word Insertions in Singlish','Multiple time formats (9:00AM, 8:45, 9:15, 2:00PM) mixed with Singlish; English words meeting/traffic/reschedule.'],

  // Type 20: Inputs with Dates (x2)
  ['Neg_0040','M','2026-05-10 walata submit karanna.','2026-05-10 වලට submit කරන්නේ.','','','Inputs with Dates; Isolated English Word Insertions in Singlish','ISO date format "2026-05-10" embedded in command; English word "submit" inserted.'],
  ['Neg_0041','M','oya dannawa da March 31 wenakota deadline eka. mama April 1st walata submit karannam kiyala hithuwa eth April 5 walata ahala.','ඔය දන්නවා ද March 31 වෙනකොට deadline එකේ. මම April 1st වලට submit කරන්නම් කියලා හිතුව ඒත් April 5 වලට ඇහලා.','','','Inputs with Dates; Question forms; Inputs with Numbers and Numeric Suffixes','Month names (March, April) + date formats (31, 1st, 5) + ordinal suffix; embedded question.'],

  // Type 21: Inputs with Unit of Measurements (x2)
  ['Neg_0042','S','kg 2k gannawa ne.','kg 2ක් ගන්නවා නේ.','','','Inputs with Unit of Measurements; Inputs with Numbers and Numeric Suffixes','Unit "kg" + number with "k" suffix. System may transliterate the measurement abbreviation.'],
  ['Neg_0043','M','mama km 12k withra yana. e gase usa metra 8k withra wewa. gedara sita km 5 withara giya.','මම km 12ක් විතරේ යනේ. ඒ ගසේ උස මෙටෘ 8ක් විතරේ වේවා. ගෙදරේ සිට km 5 විතරේ ගිය.','','','Inputs with Unit of Measurements; Inputs with Numbers and Numeric Suffixes','Multiple units (km, metra) with numeric suffixes (12k, 8k) in same input. System handles each unit inconsistently.'],

  // Type 22: Inputs with Slang and Casual Phrasing (x2)
  ['Neg_0044','S','bro mara kattak ne uba.','බ්‍රෝ මාර කට්ටක් නේ උඹ.','','','Inputs with Slang and Casual Phrasing','Slang terms: "bro", "mara" (intensifier), "kattak" (a shot/great job), "uba" (you - rude/casual). System fails on slang mapping.'],
  ['Neg_0045','M','uba nikan awe balamu kiyala. sirawata meka elakiri set eka bro. hondatama set wela thiyenawa. ape gang eke hemoma hadanawa.','උඹ නිකන් ආකෙව් බලාමු කියලා. සිරාවට මෙකා එළකිරි set එකේ බ්‍රෝ. හොඳාටම set වෙලා තිෙනවා. අපේ gang එකේ හෙමෝම හදනවා.','','','Inputs with Slang and Casual Phrasing; Isolated English Word Insertions in Singlish','Multiple slang terms (uba, nikan, sirawata, elakiri, gang) + English word set embedded.'],

  // Type 23: Online Identifiers in Singlish (x2)
  ['Neg_0046','M','mee link eka balanna: https://sinhala.wiki ne. mama eka pennanna hithuwa.','මේ link එකේ බලන්නේ: https://sinhala.wiki නේ. මම ඒකා පෙන්නන්නේ හිතුව.','','','Online Identifiers in Singlish; English Digital Terms in Singlish','URL "https://sinhala.wiki" embedded in Singlish. System must preserve URL without transliterating.'],
  ['Neg_0047','M','mama email eka ewanna: student.it23361522@university.lk kiyala danna. @Nimal oyata eka send karada?','මම email එකේ එවන්නේ: student.it23361522@university.lk කියලා දාන්නේ. @Nimal ඔයාට ඒකා send කරාද?','','','Online Identifiers in Singlish; Person Names Embedded in Singlish; Question forms','Email address + @mention identifier; person name Nimal after @; question embedded.'],

  // Type 24: Inputs Containing Emojis (x3)
  ['Neg_0048','S','honda weda karalane 😊 bro.','හොඳ වැඩේ කරලානේ 😊 බ්‍රෝ.','','','Inputs Containing Emojis; Inputs with Slang and Casual Phrasing','Emoji 😊 placed mid-sentence between Singlish words. System may strip emoji or corrupt surrounding words.'],
  ['Neg_0049','M','aiyo machan 😢 mokada wune? mama kiwwa ne denna kiyala 😡 eth oya eyata dunnada?','අයියෝ මචන් 😢 මොකද වුනේ? මම කිව්වා නේ දෙන්නේ කියලා 😡 ඒත් ඔය ඒයාට දුන්නාද?','','','Inputs Containing Emojis; Inputs with Slang and Casual Phrasing; Question forms','Two emojis (😢 😡) embedded; slang "machan"; two questions mixed with emotional context.'],
  ['Neg_0050','L','mama ada office awe passe Suresh aawa gedara. api katha kala 😊. mama kiwwa "heta 9:00AM meeting ekak thiyenawa" kiyala. Suresh kiwwa "ok machan, mama ASAP ennam" kiyala. passe api Rs. 250 wage badu kaewwa. mama kiwwa March 20 exam kiyala calendar eke 2026-03-20 damma. eka duwas hadissiye eka km 5k withra payin yana una. 😄 honda dinayak.','මම අද office ආවේ පස්සේ Suresh ආවා ගෙදර. අපි කතා කළා 😊. මම කිව්වා "හෙට 9:00AM meeting එකක් තිෙනවා" කියලා. Suresh කිව්වා "ok මචන්, මම ASAP එන්නම්" කියලා. පස්සේ අපි Rs. 250 වාගේ බඩු කෑවා. මම කිව්වා March 20 exam කියලා calendar එකේ 2026-03-20 දාම්ම. ඒකා දුවස හදිස්සියේ ඒකා km 5ක් විතරේ පයින් යන්නේ උනා. 😄 හොඳ දිනයක්.','','','Inputs Containing Emojis; Person Names Embedded in Singlish; Inputs with Time Formats; Inputs with Currency; Inputs with Dates; Inputs with Unit of Measurements; English Abbreviations/Acronyms in Singlish; Isolated English Word Insertions in Singlish; Inputs with Slang and Casual Phrasing','L-type 300-450 chars; emojis 😊😄; person name Suresh; time 9:00AM; currency Rs.250; date 2026-03-20 and March 20; unit km 5k; acronym ASAP; English words office/meeting/calendar/exam; slang machan/ok.'],
];

const headers = [
  'TC ID',
  'Input length type',
  'Input',
  'Expected output',
  'Actual output',
  'Status',
  'Singlish input types covered',
  'Evidence or rationale for the input type covered'
];

const rows = cases.map(c => c);

const wsData = [headers, ...rows];
const ws = xlsx.utils.aoa_to_sheet(wsData);

ws['!cols'] = [
  { wch: 12 },
  { wch: 18 },
  { wch: 60 },
  { wch: 60 },
  { wch: 60 },
  { wch: 10 },
  { wch: 50 },
  { wch: 80 }
];

const wb = xlsx.utils.book_new();
xlsx.utils.book_append_sheet(wb, ws, 'Test Cases');
xlsx.writeFile(wb, 'IT23361522_Assignment 1 - Test cases.xlsx');
console.log('Done: IT23361522_Assignment 1 - Test cases.xlsx created with', cases.length, 'rows.');
