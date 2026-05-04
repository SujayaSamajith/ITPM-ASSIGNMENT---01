const xlsx = require('xlsx');

const file = 'IT23361522_Assignment 1 - Test cases.xlsx';

const dataToFill = [
  ["අද school ගියනවා ද?", "Fail"],
  ["මම දන්නේ නා ඔය කොහෙද ගියේ කියලා. අපි හෙට යමුද නැත්නම් අවුරුද්ද වාගේ කරන්නෝ?", "Fail"],
  ["ඔය API එකේ documentation එකේ බලලා ද? මම එන්නේ ඕනේ නම් ticket එකක් ගන්නද?", "Fail"],
  ["වහාම ගෙදර යන්නේ! එපා ඉන්නේ.", "Fail"],
  ["ඒකා කරන්නම් කියලා කියන්නේ එපා. මම කියනවා එහෙම කරන්නේ. හදිල්ලා දන් එන්නකෝ මචන්.", "Fail"],
  ["කොහොමද බ්රෝ? සුද්ද ද ඉන්නේ?", "Fail"],
  ["මල්ලි! ගොඩක් දවසක් පස්සේ කෙතුවා නේ. කොහොමද ඉන්නේ ඔය? හොඳ වෙලා තිෙනවා වාගේ පෙනෙවා.", "Fail"],
  ["කරුණාකරා මෙකා පොඩ්ඩක් බලන්නේ පුළුවන්ද?", "Fail"],
  ["මම ඔබ කියනවා, ඔය හෙට meeting එකේදී Zoom link එකේ එවන්නේ පුළුවන්ද? මම office එකේ ඉන්නවා හදිස්සියේ.", "Fail"],
  ["අනේ හරි අයිේ. මම ඒකා කරන්නම්.", "Fail"],
  ["නැතනේ මම දන්නේ. ඒත් ඔය කියනවා නම් ඒකා කරන්නේ. මම හිතුව කරන්නේ ඒත් බැලුවේ නාහා.", "Fail"],
  ["බොරු බොරු කතා කියන්නේ එපා.", "Fail"],
  ["හරි හරි මම එන්නම්. ගොඩක් ගොඩක් වාගේ ඉන්නවා කියලානේ. පොඩ්ඩක් පොඩ්ඩක් කරාකරාලා මම ඒකා හදාගන්නම්.", "Fail"],
  ["අනේ! ඔය එනවාද... නේ?", "Fail"],
  ["මම කියනවා, ඔය දන්නා ද මෙකා? හදන්නේ නේ... ඒත් හරි කරන්නේ ඕනේ, නේ! මම ඒකා බලාමු.", "Fail"],
  ["මේ එකේ හොඳ නේ මචන්?", "Fail"],
  ["මං ඔයාට පස්සේ එන්නම් කියලා කිව්වා. ඒත් ඔයාට බලාගෙන ඉන්නේ නේ? මංගෙ ගෙදර ඇහුවොත් එනවා නම්.", "Fail"],
  ["lunch කරන්නේ යමු නේ?", "Fail"],
  ["මම අද project එකේ finish කරන්නම් කියලා හිතුව ඒත් deadline එකේ miss උනා නිසා manager කියනවා extension එකක් ගන්නේ කියලා.", "Fail"],
  ["මම කියනවා out of stock වෙලා තිෙනවා ඒත් just in case order කරන්නේ හිතුව නේ.", "Fail"],
  ["ඔය දන්නවා ද free of charge ලැබෙන බව? ඒත් terms and conditions බලන්නේ ඕනේ කියලානේ. as soon as possible submit කරන්නේ.", "Fail"],
  ["download link එකේ පෙන්නන්නේ.", "Fail"],
  ["ඔය account එකේ login කරන්නේ පුළුවන් ඒත් password reset කරන්නේ ඕනේ කියන්නේ කොච්චරද? email එකේ confirm වෙවිද?", "Fail"],
  ["Viber එකෙන් message එකක් එව.", "Fail"],
  ["මම Facebook එකේ post එකක් දාම්ම අපි Instagram වලට share කරන්නම් කියලා කිව්වා. TikTok එකේ trend එකේ බල්ලකෝ.", "Fail"],
  ["ASAP කරන්නම් ok ද?", "Fail"],
  ["ඔය CV එකේ update කරලාද? HR කියනවා ASAP submit කරන්නේ කියලා. ඒත් DOB සහ NIC එකේ හරිද check කරන්නේ.", "Fail"],
  ["prof එකෙන් lab assignment එකක් තිෙනවා.", "Fail"],
  ["අද uni එකේ lec කියලා පෙනෙලා. මම lib එකේ sit කරන්නම් කියලා හිතුව ඒත් doc එකක් print කරන්නේ ගිය.", "Fail"],
  ["Galle යමුද අද?", "Fail"],
  ["මම Nugegoda සිට Malabe යන්නේ ඕනේ. Kesbewa pass වෙනකොට call කරන්නම්. Boralesgamuwa hotel එකේ lunch කරන්නම්.", "Fail"],
  ["Ishara හෙට ආවාද?", "Fail"],
  ["Kavindra සහ Thisara දෙන්නාම project එකේ work කරලා මාරු. ඒත් Dinusha කියනවා එහෙම නේ කියලා. ඔය Pavithra කියනවා ඒකා ගැනා?", "Fail"],
  ["2nd match එකේ score එකේ කියන්නේ.", "Fail"],
  ["science lab එකේ 1st group එකේ 98% අරගෙනේ. 2nd group එකේ 85k වාගේ markings. 3rd group වලට 50 සිට 75 දාකාවා.", "Fail"],
  ["Rs. 1500 දෙන්නකෝ මචන්.", "Fail"],
  ["මම කිව්වා LKR 4500 වාගේ වේ යි කියලා. ඒත් USD 15 කියන්නේ රුපියල් වලින් කොච්චරද? EUR 100 කියන්නේ ගොඩක් නේ.", "Fail"],
  ["4:30PM නාගින් එන්නේ කරන්නේ.", "Fail"],
  ["meeting එකේ 9:00AM වලට set කරනේ. මම 8:45 වලට එන්නේ හිතුව ඒත් traffic නිසා 9:15 වලට උනා. tomorrow 2:00PM වලට reschedule කරන්නම්.", "Fail"],
  ["2026-05-10 වලට submit කරන්නේ.", "Fail"],
  ["ඔය දන්නවා ද March 31 වෙනකොට deadline එකේ. මම April 1st වලට submit කරන්නම් කියලා හිතුව ඒත් April 5 වලට ඇහලා.", "Fail"],
  ["kg 2ක් ගන්නවා නේ.", "Fail"],
  ["මම km 12ක් විතරේ යනේ. ඒ ගසේ උස මෙටෘ 8ක් විතරේ වේවා. ගෙදරේ සිට km 5 විතරේ ගිය.", "Fail"],
  ["බ්රෝ මාර කට්ටක් නේ උඹ.", "Fail"],
  ["උඹ නිකන් ආකෙව් බලාමු කියලා. සිරාවට මෙකා එළකිරි set එකේ බ්රෝ. හොඳාටම set වෙලා තිෙනවා. අපේ gang එකේ හෙමෝම හදනවා.", "Fail"],
  ["මේ link එකේ බලන්නේ: https://sinhala.wiki නේ. මම ඒකා පෙන්නන්නේ හිතුව.", "Fail"],
  ["මම email එකේ එවන්නේ: student.it23361522@university.lk කියලා දාන්නේ. @Nimal ඔයාට ඒකා send කරාද?", "Fail"],
  ["හොඳ වැඩේ කරලානේ 😊 බ්රෝ.", "Fail"],
  ["අයියෝ මචන් 😢 මොකද වුනේ? මම කිව්වා නේ දෙන්නේ කියලා 😡 ඒත් ඔය ඒයාට දුන්නාද?", "Fail"],
  ["මම අද office ආවේ පස්සේ Suresh ආවා ගෙදර. අපි කතා කළා 😊. මම කිව්වා \"හෙට 9:00AM meeting එකක් තිෙනවා\" කියලා. Suresh කිව්වා \"ok මචන්, මම ASAP එන්නම්\" කියලා. පස්සේ අපි Rs. 250 වාගේ බඩු කෑවා. මම කිව්වා March 20 exam කියලා calendar එකේ 2026-03-20 දාම්ම. ඒකා දුවස හදිස්සියේ ඒකා km 5ක් විතරේ පයින් යන්නේ උනා. 😄 හොඳ දිනයක්.", "Fail"]
];

try {
  const wb = xlsx.readFile(file);
  const wsName = wb.SheetNames[0];
  const ws = wb.Sheets[wsName];

  // The actual output is column E (index 4) and Status is column F (index 5)
  // Rows start at index 1 (row 2 in excel)
  for (let i = 0; i < dataToFill.length; i++) {
    const row = i + 1; // +1 for header row
    const actualText = dataToFill[i][0];
    const statusText = dataToFill[i][1];

    ws[xlsx.utils.encode_cell({ r: row, c: 4 })] = { t: 's', v: actualText };
    ws[xlsx.utils.encode_cell({ r: row, c: 5 })] = { t: 's', v: statusText };
  }

  // Ensure range covers the modified cells
  if (ws['!ref']) {
    const range = xlsx.utils.decode_range(ws['!ref']);
    if (range.e.c < 5) range.e.c = 5;
    if (range.e.r < dataToFill.length) range.e.r = dataToFill.length;
    ws['!ref'] = xlsx.utils.encode_range(range);
  }

  xlsx.writeFile(wb, file);
  console.log('Successfully updated Excel file without touching other columns.');
} catch (err) {
  console.error('Error:', err);
}
