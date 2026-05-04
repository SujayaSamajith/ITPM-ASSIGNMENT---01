# IT23361522 – Assignment 1: Transliteration Accuracy Testing

**Student:** IT23361522  
**Module:** IT3040 – IT Project Management  
**Assignment:** Assignment 1 – Option 1: Transliteration Accuracy Testing  
**Target Website:** https://www.pixelssuite.com/chat-translator  

---

## Project Overview

This project contains **50 negative test cases** that test the **Chat Sinhala** transliteration function on PixelsSuite.  
A *negative test case* is one where the system is expected to **fail** – i.e., the actual transliterated output is different from the correct expected Sinhala output.

All 24 Singlish input types are covered, with at least 2 test cases per type.  
Test IDs run from `Neg_0001` to `Neg_0050`.

### Files in This Project

| File | Purpose |
|------|---------|
| `IT23361522_Assignment 1 - Test cases.xlsx` | 50 negative test cases (auto-updated by the test script) |
| `tests/translator.spec.js` | Playwright automation script |
| `playwright.config.js` | Playwright configuration |
| `create-excel.js` | One-time script to regenerate the Excel file from scratch |
| `inspect.js` | Helper script to dump the raw page HTML for selector debugging |
| `package.json` | Project dependencies and `npm test` command |

---

## Prerequisites

- [Node.js](https://nodejs.org/) v16 or higher  
- npm (comes with Node.js)

---

## Installation

```bash
# 1. Navigate to the project folder
cd "IT23361522"

# 2. Install Node dependencies
npm install

# 3. Install the Playwright Chromium browser
npx playwright install chromium
```

---

## Running the Tests

```bash
npm test
```

This runs `playwright test` which will:
1. Read all 50 test cases from the Excel file
2. Open https://www.pixelssuite.com/chat-translator in Chrome
3. Type each Singlish input into the input box
4. Wait for the transliteration output
5. Compare it with the expected Sinhala
6. Write the captured output and Pass/Fail status back into the Excel file
7. Save the updated Excel

---

## Viewing the HTML Test Report

After the tests finish, open the Playwright HTML report:

```bash
npx playwright show-report
```

This opens a detailed report in your browser showing each test, its status, and any errors.

---

## Regenerating the Excel File

If you need to recreate the Excel from scratch (e.g., to reset Actual output and Status columns):

```bash
node create-excel.js
```

---

## Notes

### Selectors
The test script uses `page.locator('textarea').first()` for input and `page.locator('textarea').nth(1)` for output.  
If the website DOM changes, update these lines in `tests/translator.spec.js`:

```js
const inputBox  = page.locator('textarea').first();   // line ~39
const outputBox = page.locator('textarea').nth(1);    // line ~40
```

Alternative selectors (commented out in the script):
```js
const inputBox  = page.getByPlaceholder('Type your English text here');
const outputBox = page.getByPlaceholder('Transliterated Sinhala');
```

To find the correct selectors, run the inspect helper:
```bash
node inspect.js
```
Then search the printed HTML for `textarea` or `placeholder` attributes.

### Input Length Types
| Code | Range |
|------|-------|
| S | 30 characters or fewer |
| M | 31 – 299 characters |
| L | 300 – 450 characters |

### Test Status
| Status | Meaning |
|--------|---------|
| Fail | Actual output ≠ Expected output (expected for negative cases) |
| Pass | Actual output = Expected output (should not happen for negative cases) |
