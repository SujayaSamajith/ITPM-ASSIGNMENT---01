# Instructions

- Following Playwright test failed.
- Explain why, be concise, respect Playwright best practices.
- Provide a snippet of code with the fix, if possible.

# Test info

- Name: translator.spec.js >> IT23361522 – Singlish to Sinhala Negative Test Cases >> Neg_0042 | Inputs with Unit of Measurements
- Location: tests\translator.spec.js:40:5

# Error details

```
Error: Neg_0042: system should fail on this input (negative case)

expect(received).not.toEqual(expected) // deep equality

Expected: not "kg 2ක් ගන්නවා නේ."

```

# Test source

```ts
  3   | // IT3040 ITPM – Option 1 (Negative Test Cases)
  4   | //
  5   | // This script:
  6   | //  1. Reads TC ID, Input, Expected output from the Excel file
  7   | //  2. Opens https://www.pixelssuite.com/chat-translator
  8   | //  3. Types each Singlish input into the input box
  9   | //  4. Waits for the Chat Sinhala transliteration output
  10  | //  5. Compares actual vs expected output
  11  | //  6. Writes Actual output + Status back into the Excel file
  12  | // ============================================================
  13  | 
  14  | const { test, expect } = require('@playwright/test');
  15  | const xlsx = require('xlsx');
  16  | const path = require('path');
  17  | 
  18  | // ── Excel file path ────────────────────────────────────────────────────────────
  19  | const EXCEL_FILE = path.join(
  20  |   __dirname, '..', 'IT23361522_Assignment 1 - Test cases.xlsx'
  21  | );
  22  | 
  23  | // ── Load all test cases from Excel ────────────────────────────────────────────
  24  | const workbook  = xlsx.readFile(EXCEL_FILE);
  25  | const sheetName = workbook.SheetNames[0];
  26  | const sheet     = workbook.Sheets[sheetName];
  27  | const records   = xlsx.utils.sheet_to_json(sheet);
  28  | 
  29  | // ── Shared results collector (written back in afterAll) ────────────────────
  30  | const results = [];
  31  | 
  32  | // Run tests one by one so the afterAll Excel write is safe
  33  | test.describe.configure({ mode: 'serial' });
  34  | 
  35  | test.describe('IT23361522 – Singlish to Sinhala Negative Test Cases', () => {
  36  | 
  37  |   for (let i = 0; i < records.length; i++) {
  38  |     const record = records[i];
  39  | 
  40  |     test(`${record['TC ID']} | ${record['Singlish input types covered'].split(';')[0].trim()}`, async ({ page }) => {
  41  | 
  42  |       // ── 1. Navigate to the Chat Translator ──────────────────────────────
  43  |       try {
  44  |         await page.goto('https://www.pixelssuite.com/chat-translator', {
  45  |           waitUntil: 'domcontentloaded',
  46  |           timeout: 30000,
  47  |         });
  48  |       } catch {
  49  |         test.skip(true, 'Website is unreachable.');
  50  |         return;
  51  |       }
  52  | 
  53  |       // ── 2. Wait for the page to fully render ────────────────────────────
  54  |       await page.waitForTimeout(3000);
  55  | 
  56  |       // ── UI Selectors ─────────────────────────────────────────────────────
  57  |       // NOTE: If the website DOM changes, update these selectors.
  58  |       // Run `node inspect.js` to dump the live page HTML and find correct values.
  59  |       const inputBox  = page.locator('textarea').first();   // Singlish input area
  60  |       const outputBox = page.locator('textarea').nth(1);    // Sinhala output area
  61  | 
  62  |       // Alternative selectors (uncomment if above don't work):
  63  |       // const inputBox  = page.getByPlaceholder('Type your English text here');
  64  |       // const outputBox = page.getByPlaceholder('Transliterated Sinhala');
  65  | 
  66  |       // ── 3. Type the Singlish input ───────────────────────────────────────
  67  |       await inputBox.click();
  68  |       await inputBox.fill('');
  69  |       await inputBox.fill(String(record['Input']));
  70  | 
  71  |       // Click the transliterate button
  72  |       await page.locator('button:has-text("Transliterate")').click();
  73  | 
  74  |       // ── 4. Wait for transliteration to process ──────────────────────────
  75  |       await page.waitForTimeout(3000);
  76  | 
  77  |       // ── 5. Capture the actual output ─────────────────────────────────────
  78  |       let actualOutput = '';
  79  |       try {
  80  |         actualOutput = await outputBox.inputValue();
  81  |         if (!actualOutput) actualOutput = await outputBox.innerText();
  82  |       } catch {
  83  |         actualOutput = 'ERROR: Could not read output';
  84  |       }
  85  | 
  86  |       const expectedOutput = String(record['Expected output']);
  87  | 
  88  |       // ── 6. Determine Pass / Fail ─────────────────────────────────────────
  89  |       // Negative test cases FAIL when actual output ≠ expected (correct) output.
  90  |       const status = (actualOutput.trim() !== expectedOutput.trim()) ? 'Fail' : 'Pass';
  91  | 
  92  |       // ── 7. Save result for Excel write-back ──────────────────────────────
  93  |       results.push({ rowIndex: i, actualOutput, status });
  94  | 
  95  |       console.log(`[${record['TC ID']}] Status: ${status}`);
  96  |       console.log(`  Expected: "${expectedOutput.substring(0, 60)}"`);
  97  |       console.log(`  Actual  : "${actualOutput.substring(0, 60)}"\n`);
  98  | 
  99  |       // ── 8. Assert – negative cases should produce wrong output ──────────
  100 |       expect(
  101 |         actualOutput.trim(),
  102 |         `${record['TC ID']}: system should fail on this input (negative case)`
> 103 |       ).not.toEqual(expectedOutput.trim());
      |             ^ Error: Neg_0042: system should fail on this input (negative case)
  104 |     });
  105 |   }
  106 | 
  107 |   // ── After all tests: write Actual output + Status back to Excel ──────────
  108 |   test.afterAll(async () => {
  109 |     try {
  110 |       const wb = xlsx.readFile(EXCEL_FILE);
  111 |       const ws = wb.Sheets[wb.SheetNames[0]];
  112 | 
  113 |       const rawData   = xlsx.utils.sheet_to_json(ws, { header: 1 });
  114 |       const headerRow = rawData[0];
  115 |       const colActual = headerRow.indexOf('Actual output');
  116 |       const colStatus = headerRow.indexOf('Status');
  117 | 
  118 |       if (colActual === -1 || colStatus === -1) {
  119 |         console.error('ERROR: Could not find "Actual output" or "Status" columns.');
  120 |         return;
  121 |       }
  122 | 
  123 |       for (const result of results) {
  124 |         const row = result.rowIndex + 1; // +1 for header row
  125 | 
  126 |         ws[xlsx.utils.encode_cell({ r: row, c: colActual })] =
  127 |           { t: 's', v: result.actualOutput };
  128 |         ws[xlsx.utils.encode_cell({ r: row, c: colStatus })] =
  129 |           { t: 's', v: result.status };
  130 |       }
  131 | 
  132 |       xlsx.writeFile(wb, EXCEL_FILE);
  133 |       console.log('✅ Excel updated:', EXCEL_FILE);
  134 |     } catch (err) {
  135 |       console.error('Failed to write results back to Excel:', err.message);
  136 |     }
  137 |   });
  138 | });
  139 | 
```