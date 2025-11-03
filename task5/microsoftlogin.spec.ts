import { test, expect } from "@playwright/test";
import dotenv from "dotenv";
import { authenticator } from "otplib";
test.setTimeout(120000); // 2 minutes total timeout


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

  // Step 10: Wait for Office dashboard to load (more flexible)
await page.waitForLoadState("domcontentloaded", { timeout: 30000 });

// Try to detect dashboard UI element
const dashboardReady = await page
  .locator('a:has-text("Word"), a[aria-label*="Word"]')
  .first()
  .isVisible()
  .catch(() => false);

if (!dashboardReady) {
  console.warn("⚠️ Dashboard still loading... adding delay");
  await page.waitForTimeout(10000);
}

console.log("📊 Office dashboard loaded (relaxed mode)!");


// Step 11: Locate Word app link (robust locator)
console.log("📝 Searching for Word app...");

await page.waitForTimeout(8000); // Give dashboard time to render
let wordLink = page.locator('a:has-text("Word"), a[aria-label*="Word"]');

// Retry logic: check visibility or fallback text
if (!(await wordLink.first().isVisible())) {
  console.log("⚠️ Word link not visible yet, retrying...");
  await page.waitForTimeout(10000);

  // Broaden search text variations
  wordLink = page.locator('a:has-text("Word"), a:has-text("Word Online"), [aria-label*="Word"], [title*="Word"]');
}

await wordLink.first().waitFor({ state: "visible", timeout: 60000 });
console.log("✅ Word app found — clicking...");

// Handle new tab opening
const [newTabe] = await Promise.all([
  page.context().waitForEvent("page").catch(() => null),
  wordLink.first().click(),
]);

const wordHandlee = newTabe || page;
await wordHandlee.waitForLoadState("domcontentloaded", { timeout: 60000 });
console.log("🌐 Word page opened successfully:", wordHandlee.url());


// Step 12: Handle new tab or fallback to same page
const [maybeNewTab] = await Promise.all([
  page.context().waitForEvent("page").catch(() => null),
  wordLink.first().click(),
]);

let wordHandle = maybeNewTab;
if (!wordHandle || wordHandle.isClosed()) {
  console.log("⚠️ No new tab detected — using main page instead");
  wordHandle = page;
}

// Wait for Word to load safely
try {
  await wordHandle.waitForLoadState("domcontentloaded", { timeout: 60000 });
  await wordHandle.waitForTimeout(5000);
  console.log("🌐 Word page loaded successfully:", wordHandle.url());
} catch (err) {
  console.error("❌ Word page did not load correctly:", err.message);
  await page.screenshot({ path: "word_error.png", fullPage: true });
}


// Step 13: Get Word tab handle safely
const wordHandles = newTabe || page;
await wordHandles.waitForLoadState("domcontentloaded");
await wordHandles.waitForTimeout(20000);
console.log("🌐 Word page loaded:", wordHandles.url());

// Optional: verify Word domain
const wordURL = wordHandles.url();
expect(wordURL).toMatch(/(word\.cloud\.microsoft|word\.office\.com|office\.live\.com)/);

// Step 14: Create a new blank document
console.log("📝 Creating a new blank document...");
const newBlankButton =
  (await wordHandles.$('text="New blank document"')) ||
  (await wordHandles.$('button:has-text("New blank document")')) ||
  (await wordHandles.$('[aria-label="New blank document"]'));

if (newBlankButton) {
  await newBlankButton.click();
  console.log("✅ Clicked 'New blank document' successfully!");
} else {
  console.warn("⚠️ Could not find the 'New blank document' button!");
}

await wordHandle.waitForTimeout(5000);
console.log("🎉 Word Online launched successfully!");

});

















