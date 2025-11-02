import { test, expect } from "@playwright/test";
import dotenv from "dotenv";
import { authenticator } from "otplib";

dotenv.config();

const EMAIL = process.env.MS_EMAIL!;
const PASSWORD = process.env.MS_PASSWORD!;
const TOTP_SECRET = process.env.MS_TOTP_SECRET!;

test("Microsoft Office login", async ({ page }) => {
  console.log("🚀 Starting Microsoft login automation...");

  // Step 1: Go to Office
  await page.goto("https://www.office.com/");
  await page.getByRole("link", { name: /Sign in/i }).click();

  // Step 2: Email
  const emailInput = page.getByRole("textbox", { name: /Email|Sign in/i });
  await emailInput.waitFor({ state: "visible", timeout: 40000 });
  await emailInput.fill(EMAIL);
  await page.getByRole("button", { name: /Next/i }).click();

  // Step 3: Password
  const passwordInput = page.getByRole("textbox", { name: /Password/i });
  await passwordInput.waitFor({ state: "visible", timeout: 40000 });
  await passwordInput.fill(PASSWORD);
  await page.getByRole("button", { name: /^Sign in$|^Next$/i }).click();

  // Step 4: MFA setup (use verification code)
  const useCodeBtn = page.getByRole("button", { name: /Use a verification code/i });
  if (await useCodeBtn.isVisible({ timeout: 10000 }).catch(() => false)) {
    console.log("👉 Clicking 'Use a verification code'");
    await useCodeBtn.click();
  }

  const haveCodeBtn = page.getByRole("button", { name: /I have a code/i });
  if (await haveCodeBtn.isVisible({ timeout: 10000 }).catch(() => false)) {
    console.log("👉 Clicking 'I have a code'");
    await haveCodeBtn.click();
  }

  // Step 5: Wait for OTP input (Authenticator app code)
  const otpInput = page.getByRole("textbox", { name: /^Code$/i });
  await otpInput.waitFor({ state: "visible", timeout: 40000 });

  // Step 6: Generate & fill OTP
  authenticator.options = { step: 30, window: 1 };
  let otp = authenticator.generate(TOTP_SECRET);
  console.log("🔢 Generated OTP:", otp);

  await otpInput.fill(otp);

  // The button text is “Next” not “Verify” on this page
  const nextButton = page.getByRole("button", { name: /^Next$/i });
  await nextButton.waitFor({ state: "visible", timeout: 10000 });     
  await nextButton.click({ force: true });
  console.log("✅ Clicked Next after entering OTP");

  // Step 7: Retry if OTP fails
  const errorAlert = page.getByText(/incorrect|didn't work|try again/i);
  if (await errorAlert.isVisible({ timeout: 4000 }).catch(() => false)) {
    console.log("⚠️ First OTP failed — retrying...");
    otp = authenticator.generate(TOTP_SECRET);
    await otpInput.fill(otp);
    await nextButton.click({ force: true });
  }


  // Step 8: Handle “Stay signed in?” — supports iframe and main page
  await page.waitForTimeout(2000);
  console.log("🟢 Checking for 'Stay signed in?' screen...");

  // Find frame that contains the heading
  const frames = page.frames();
  let stayFrame = frames.find(f =>
    f.url().includes("login.live.com") || f.url().includes("ppsecure")
  );

  // Try current page if no frame found
  const context = stayFrame || page;

  const stayHeading = context.getByRole("heading", { name: /Stay signed in\?/i });
  if (await stayHeading.isVisible({ timeout: 15000 }).catch(() => false)) {
    console.log("🟢 'Stay signed in?' detected");

    // Choose action
    const yesButton = context.getByRole("button", { name: /^Yes$/i });
    const noButton = context.getByRole("button", { name: /^No$/i });

    // 👇 Change this line to yesButton if you prefer staying logged in
    const targetButton = noButton;

    if (await targetButton.isVisible({ timeout: 8000 }).catch(() => false)) {
      await targetButton.click({ force: true });
      console.log(`✅ Clicked '${targetButton === yesButton ? "Yes" : "No"}' on Stay signed in`);
    } else {
      console.log("⚠️ Buttons not clickable — skipping step.");
    }
  } else {
    console.log("⏭️ No Stay signed in screen detected.");
  }

  // Step 9: Wait for dashboard
  await page.waitForLoadState("networkidle", { timeout: 30000 });
  await expect(page).toHaveURL(/m365\.cloud\.microsoft\/search/);
  console.log("🎉 Login successful — reached Office dashboard!");

  // Step 10–11: Open Word app (handles portal variations)
  console.log("📄 Searching for 'Word' app...");

  // Find all frames that might contain app shortcuts
  const allFrames = page.frames();

  // Try multiple locator patterns
  const wordLocators = [
    page.locator('a:has-text("Word")'),
    page.locator('button:has-text("Word")'),
    ...allFrames.map(f => f.locator('a:has-text("Word")')),
    ...allFrames.map(f => f.locator('button:has-text("Word")'))
  ];

  // Try to find and click the first visible one
  let wordFound = false;
  for (const loc of wordLocators) {
    if (await loc.isVisible({ timeout: 5000 }).catch(() => false)) {
      console.log("✅ Found 'Word' button/link — clicking...");
      await loc.click({ force: true });
      wordFound = true;
      break;
    }
  }

  if (!wordFound) {
    throw new Error("❌ Could not find 'Word' link/button anywhere.");
  }

  // Step 11: Handle Word Online new tab
console.log("🕓 Waiting for Word Online tab to open...");

// When the Word link is clicked, a new page opens — wait for it
const [wordPage] = await Promise.all([
  page.context().waitForEvent("page"),
  // Click on Word (works for button or link)
  page.locator('a:has-text("Word"), button:has-text("Word")').first().click({ force: true })
]);

await wordPage.waitForLoadState("domcontentloaded", { timeout: 60000 });
const wordURL = wordPage.url();
console.log("🌐 Word Online opened:", wordURL);

// Check we reached a valid Word environment
expect(wordURL).toMatch(/(word\.cloud\.microsoft|word\.office\.com|office\.live\.com|m365\.cloud\.microsoft)/);
console.log("✅ Word Online tab loaded successfully!");

// Step 12: Create a new blank Word document
console.log("📄 Creating a new blank Word document...");

const newBlankDoc = wordPage.getByRole("link", { name: /New blank document|Blank document/i });
const newBlankButton = wordPage.getByRole("button", { name: /New blank document|Blank document/i });
const altBlank = wordPage.locator('a[title*="Blank document"], div:has-text("New blank document"), div:has-text("Blank document")');

await Promise.any([
  newBlankDoc.waitFor({ state: "visible", timeout: 30000 }).catch(() => null),
  newBlankButton.waitFor({ state: "visible", timeout: 30000 }).catch(() => null),
  altBlank.first().waitFor({ state: "visible", timeout: 30000 }).catch(() => null),
]);

if (await newBlankDoc.isVisible().catch(() => false)) {
  await newBlankDoc.click();
} else if (await newBlankButton.isVisible().catch(() => false)) {
  await newBlankButton.click();
} else {
  await altBlank.first().click();
}

console.log("✅ Clicked 'New blank document' successfully!");

// Step 13: Wait for editor to load
await wordPage.waitForLoadState("networkidle", { timeout: 60000 });
await expect(wordPage).toHaveURL(/word\.(cloud\.microsoft|office\.com).*document/);
console.log("📝 Blank Word document created successfully!");

});


















// import { test, expect } from "@playwright/test";
// import dotenv from "dotenv";
// import { authenticator } from "otplib";

// dotenv.config();

// const EMAIL = process.env.MS_EMAIL!;
// const PASSWORD = process.env.MS_PASSWORD!;
// const TOTP_SECRET = process.env.MS_TOTP_SECRET || "";

// test("Microsoft Office login with TOTP", async ({ page }) => {
//   await page.goto("https://www.office.com/");

//   // Click Sign in
//   await page.getByRole("link", { name: /Sign in/i }).click();

//   // Fill email
//   const emailInput = page.getByRole("textbox", { name: /Email|Sign in|Enter your email/i });
//   await emailInput.waitFor({ state: "visible", timeout: 20000 });
//   await emailInput.fill(EMAIL);
//   await page.locator('input[type="submit"], button[type="submit"], button:has-text("Next")').first().click();

//   // Wait for password page & fill password
//   const passwordBox = page.getByRole("textbox", { name: /Password/i }).first();
//   await passwordBox.waitFor({ state: "visible", timeout: 20000 });
//   await passwordBox.click();
//   await passwordBox.fill(PASSWORD);
//   await page.getByRole("button", { name: /^Next$|^Sign in$|^Submit$/i }).first().click();

//   // Wait for possible MFA / 2FA page to show up
//   // We'll try multiple selectors so it's robust across variations.
//   const totpSelectors = [
//     'input[name="otc"]',                           // common
//     'input[type="tel"]',                           // sometimes used
//     'input[aria-label*="code"]',                   // "Enter code" etc
//     'input[aria-label*="Authenticator"]',
//     'input[placeholder*="code"]',
//     'input[id*="otc"]',
//   ];

//   // Wait until either we're redirected to dashboard OR one of totp inputs appears
//   const dashboardOrOtp = await Promise.race([
//     page.waitForURL(/.*(office\.com|microsoft\.com).*$/, { timeout: 20000 }).then(() => "dashboard").catch(() => null),
//     (async () => {
//       for (const sel of totpSelectors) {
//         try {
//           const locator = page.locator(sel);
//           await locator.waitFor({ state: "visible", timeout: 7000 });
//           return sel;
//         } catch (e) {
//           // try next
//         }
//       }
//       return null;
//     })(),
//   ]);

//   if (dashboardOrOtp === "dashboard") {
//     console.log("✅ Logged in without MFA!");
//     return;
//   }

//   // If we detected an OTP input selector, fill TOTP
//   let filled = false;
//   const totp = TOTP_SECRET ? authenticator.generate(TOTP_SECRET) : null;

//   if (!totp) {
//     throw new Error("TOTP secret not found in MS_TOTP_SECRET env variable.");
//   }

//   // Try known selectors first
//   for (const sel of totpSelectors) {
//     const locator = page.locator(sel);
//     if (await locator.count() > 0) {
//       try {
//         // If single input exists, fill it
//         await locator.first().click();
//         await locator.first().fill(totp);
//         filled = true;
//         break;
//       } catch (e) {
//         // ignore and fallback
//       }
//     }
//   }

//   // Fallback: some pages use 6 separate inputs for each digit
//   if (!filled) {
//     const multiInputs = page.locator('input').filter({ has: page.locator('[inputmode="numeric"],[aria-label*="digit"],[aria-label*="code"]') });
//     const count = await multiInputs.count();
//     if (count >= 4 && count <= 8) {
//       // fill digit by digit
//       for (let i = 0; i < Math.min(totp.length, count); i++) {
//         await multiInputs.nth(i).click();
//         await multiInputs.nth(i).fill(totp[i]);
//       }
//       filled = true;
//     } else {
//       // Last resort: try any visible numeric input
//       const anyVisible = page.locator('input[type="text"], input[type="tel"], input[type="number"]').filter({ hasText: "" });
//       if ((await anyVisible.count()) > 0) {
//         await anyVisible.first().click();
//         await anyVisible.first().fill(totp);
//         filled = true;
//       }
//     }
//   }

//   if (!filled) {
//     throw new Error("Could not find OTP input to fill.");
//   }

//   // Click verify/submit
//   const submitButtons = [
//     page.getByRole("button", { name: /Verify|Next|Submit|Sign in|Confirm/i }),
//     page.locator('input[type="submit"]'),
//     page.locator('button:has-text("Verify")'),
//   ];
//   for (const btn of submitButtons) {
//     try {
//       if ((await btn.count()) > 0) {
//         await btn.first().click();
//         break;
//       }
//     } catch (e) {
//       // continue
//     }
//   }

//   // Wait for dashboard or success navigation
//   await page.waitForLoadState("networkidle");
//   await expect(page).toHaveURL(/office\.com|microsoft\.com/);

//   console.log("✅ Logged in with TOTP successfully!");
// });