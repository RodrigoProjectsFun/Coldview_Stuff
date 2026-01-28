import asyncio
import csv
from datetime import datetime
from playwright.async_api import async_playwright

# --- CONFIGURATION ---
TARGET_URL = "https://example.com"    # <--- Your URL
ELEMENT_SELECTOR = "#price-value"     # <--- Your Selector
CSV_FILENAME = "realtime_log.csv"
# ---------------------

async def run_realtime_monitor():
    # Setup CSV with headers
    with open(CSV_FILENAME, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerow(["Timestamp", "Value"])

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False) # Headless=False so you can see it working
        page = await browser.new_page()

        print(f"--- Listening for changes on: {TARGET_URL} ---")
        await page.goto(TARGET_URL)
        
        # Ensure element exists before we attach the listener
        await page.wait_for_selector(ELEMENT_SELECTOR)

        # 1. DEFINE PYTHON FUNCTION TO HANDLE CHANGES
        # This function will be called by the browser whenever it sees a change
        async def handle_change(new_value):
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3] # High precision time
            print(f"[{timestamp}] UPDATE DETECTED: {new_value}")
            
            with open(CSV_FILENAME, "a", newline="", encoding="utf-8") as f:
                csv.writer(f).writerow([timestamp, new_value])

        # 2. EXPOSE PYTHON FUNCTION TO BROWSER JAVASCRIPT
        await page.expose_function("python_on_change", handle_change)

        # 3. INJECT JAVASCRIPT OBSERVER (The "Spy")
        # This JS runs inside the Chrome browser. It watches the DOM.
        await page.evaluate(f"""(selector) => {{
            const targetNode = document.querySelector(selector);
            
            // Configuration: watch for text changes inside the element
            const config = {{ characterData: true, attributes: false, childList: true, subtree: true }};

            // Callback function to execute when mutations are observed
            const callback = (mutationList, observer) => {{
                const text = targetNode.innerText;
                // Send the new text back to Python
                window.python_on_change(text);
            }};

            // Create an observer instance linked to the callback function
            const observer = new MutationObserver(callback);

            // Start observing
            observer.observe(targetNode, config);
        }}""", ELEMENT_SELECTOR)

        print("Observer attached. Waiting for website updates...")

        # Keep script running forever
        await asyncio.Future() 

if __name__ == "__main__":
    try:
        asyncio.run(run_realtime_monitor())
    except KeyboardInterrupt:
        pass
