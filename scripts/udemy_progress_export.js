/**
 * scripts/udemy_progress_export.js
 * =================================
 * 1-Click Browser Progress Exporter for Udemy.
 *
 * HOW TO USE:
 * 1. Open your course on Udemy in your web browser:
 *    https://www.udemy.com/course/excel-vba-and-macros-course/learn/lecture/...
 * 2. Open Developer Tools (Press F12 or Ctrl+Shift+I) -> click the "Console" tab.
 * 3. Paste this entire code block into the Console and press Enter.
 * 4. A file named 'udemy_progress.json' will automatically download.
 * 5. Move the downloaded 'udemy_progress.json' into your repository root and run:
 *    python scripts/udemy_progress_sync.py --from-json udemy_progress.json
 */

(function exportUdemyProgress() {
    console.log("🔍 Scanning Udemy course curriculum for completed lectures...");
    
    // Select curriculum items
    const rows = document.querySelectorAll('li[data-purpose*="item"], div[class*="curriculum-item-link"], div[class*="item-link"]');
    const completedLectures = [];
    
    // Helper to find titles and checkmarks
    document.querySelectorAll('input[type="checkbox"]').forEach(chk => {
        if (chk.checked) {
            // Find parent row / item title
            let parent = chk.closest('li') || chk.closest('div[class*="section--item"]');
            if (parent) {
                let titleEl = parent.querySelector('span[class*="title"], div[class*="title"], span[data-purpose*="item-title"]');
                let title = titleEl ? titleEl.innerText.trim() : null;
                if (title) {
                    completedLectures.push({
                        title: title,
                        completed: true,
                        timestamp: new Date().toISOString()
                    });
                }
            }
        }
    });

    // Fallback: Check Udemy's internal React / Redux state if available
    if (completedLectures.length === 0 && window.__INITIAL_STATE__) {
        try {
            const courseState = window.__INITIAL_STATE__.courseTaking;
            if (courseState && courseState.completedItemIds) {
                completedLectures.push(...courseState.completedItemIds.map(id => ({ id: String(id), completed: true })));
            }
        } catch (e) {
            console.warn("Could not read React state directly:", e);
        }
    }

    if (completedLectures.length === 0) {
        alert("⚠️ No completed lectures detected on current view. Please ensure the Course Content sidebar is expanded!");
        return;
    }

    console.log(`✅ Found ${completedLectures.length} completed lectures! Downloading udemy_progress.json...`);

    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(completedLectures, null, 2));
    const downloadAnchor = document.createElement('a');
    downloadAnchor.setAttribute("href", dataStr);
    downloadAnchor.setAttribute("download", "udemy_progress.json");
    document.body.appendChild(downloadAnchor);
    downloadAnchor.click();
    downloadAnchor.remove();

    console.log("🎉 udemy_progress.json downloaded successfully! Run 'python scripts/udemy_progress_sync.py --from-json udemy_progress.json' to update your vault.");
})();
