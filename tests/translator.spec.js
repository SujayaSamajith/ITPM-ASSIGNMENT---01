// ============================================================
// IT23361522 – Assignment 1: Transliteration Accuracy Testing
// IT3040 ITPM – Option 1 (Negative Test Cases)
//
// This script:
//  1. Reads TC ID, Input, Expected output from the Excel file
//  2. Opens https://www.pixelssuite.com/chat-translator
//  3. Types each Singlish input into the input box
//  4. Waits for the Chat Sinhala transliteration output
//  5. Compares actual vs expected output
//  6. Writes Actual output + Status back into the Excel file
// ============================================================

const { test, expect } = require('@playwright/test');
const xlsx = require('xlsx');
const path = require('path');

// ── Excel file path ────────────────────────────────────────────────────────────
const EXCEL_FILE = path.join(
  __dirname, '..', 'IT23361522_Assignment 1 - Test cases.xlsx'
);

// ── Load all test cases from Excel ────────────────────────────────────────────
const workbook  = xlsx.readFile(EXCEL_FILE);
const sheetName = workbook.SheetNames[0];
const sheet     = workbook.Sheets[sheetName];
const records   = xlsx.utils.sheet_to_json(sheet);

// ── Shared results collector (written back in afterAll) ────────────────────
const results = [];

// Run tests one by one so the afterAll Excel write is safe
test.describe.configure({ mode: 'serial' });

test.describe('IT23361522 – Singlish to Sinhala Negative Test Cases', () => {

  for (let i = 0; i < records.length; i++) {
    const record = records[i];

    test(`${record['TC ID']} | ${record['Singlish input types covered'].split(';')[0].trim()}`, async ({ page }) => {

      // ── 1. Navigate to the Chat Translator ──────────────────────────────
      try {
        await page.goto('https://www.pixelssuite.com/chat-translator', {
          waitUntil: 'domcontentloaded',
          timeout: 30000,
        });
      } catch {
        test.skip(true, 'Website is unreachable.');
        return;
      }

      // ── 2. Wait for the page to fully render ────────────────────────────
      await page.waitForTimeout(3000);

      // ── UI Selectors ─────────────────────────────────────────────────────
      // NOTE: If the website DOM changes, update these selectors.
      // Run `node inspect.js` to dump the live page HTML and find correct values.
      const inputBox  = page.locator('textarea').first();   // Singlish input area
      const outputBox = page.locator('textarea').nth(1);    // Sinhala output area

      // Alternative selectors (uncomment if above don't work):
      // const inputBox  = page.getByPlaceholder('Type your English text here');
      // const outputBox = page.getByPlaceholder('Transliterated Sinhala');

      // ── 3. Type the Singlish input ───────────────────────────────────────
      await inputBox.click();
      await inputBox.fill('');
      await inputBox.fill(String(record['Input']));

      // Click the transliterate button
      await page.locator('button:has-text("Transliterate")').click();

      // ── 4. Wait for transliteration to process ──────────────────────────
      await page.waitForTimeout(3000);

      // ── 5. Capture the actual output ─────────────────────────────────────
      let actualOutput = '';
      try {
        actualOutput = await outputBox.inputValue();
        if (!actualOutput) actualOutput = await outputBox.innerText();
      } catch {
        actualOutput = 'ERROR: Could not read output';
      }

      const expectedOutput = String(record['Expected output']);

      // ── 6. Determine Pass / Fail ─────────────────────────────────────────
      // Negative test cases FAIL when actual output ≠ expected (correct) output.
      const status = (actualOutput.trim() !== expectedOutput.trim()) ? 'Fail' : 'Pass';

      // ── 7. Save result for Excel write-back ──────────────────────────────
      results.push({ rowIndex: i, actualOutput, status });

      console.log(`[${record['TC ID']}] Status: ${status}`);
      console.log(`  Expected: "${expectedOutput.substring(0, 60)}"`);
      console.log(`  Actual  : "${actualOutput.substring(0, 60)}"\n`);

      // ── 8. Assert – negative cases should produce wrong output ──────────
      expect(
        actualOutput.trim(),
        `${record['TC ID']}: system should fail on this input (negative case)`
      ).not.toEqual(expectedOutput.trim());
    });
  }

  // ── After all tests: write Actual output + Status back to Excel ──────────
  test.afterAll(async () => {
    try {
      const wb = xlsx.readFile(EXCEL_FILE);
      const ws = wb.Sheets[wb.SheetNames[0]];

      const rawData   = xlsx.utils.sheet_to_json(ws, { header: 1 });
      const headerRow = rawData[0];
      const colActual = headerRow.indexOf('Actual output');
      const colStatus = headerRow.indexOf('Status');

      if (colActual === -1 || colStatus === -1) {
        console.error('ERROR: Could not find "Actual output" or "Status" columns.');
        return;
      }

      for (const result of results) {
        const row = result.rowIndex + 1; // +1 for header row

        ws[xlsx.utils.encode_cell({ r: row, c: colActual })] =
          { t: 's', v: result.actualOutput };
        ws[xlsx.utils.encode_cell({ r: row, c: colStatus })] =
          { t: 's', v: result.status };
      }

      xlsx.writeFile(wb, EXCEL_FILE);
      console.log('✅ Excel updated:', EXCEL_FILE);
    } catch (err) {
      console.error('Failed to write results back to Excel:', err.message);
    }
  });
});
