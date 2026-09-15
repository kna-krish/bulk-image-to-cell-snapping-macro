# Excel Image Layout Automation Script (Self-Healing Grid Alignment Utility)

A powerful, robust VBA macro designed to fix Excel's native layout lag when extracting embedded cell images (`Picture in Cell`) and converting them to floating layouts (`Place Over Cells`). 

This script features a **self-healing loop** and an **automated diagnostic logging system** to ensure images never get skipped, misaligned, or stuck in background threads.

## 🚀 Key Features
- **Interactive UI Prompts:** Prompts the user for starting rows, ending rows, column targets, row heights, and column widths. No coding required to use it!
- **Self-Healing Mechanics:** Detects background rendering failures (like Excel skipping even rows) and automatically uses a clipboard bypass (`CopyPicture`) to ensure 100% conversion success.
- **Smart Aspect-Ratio Fitting:** Automatically rescales and centers images perfectly inside cell grid boundaries without distorting or stretching the image.
- **Auto-Refresh Engine:** Toggles Excel's display view behind the scenes to force a graphics redraw, instantly snapping all floating images into focus.
- **Automated Logging:** Saves a detailed text log with a precise timestamp (`excel_centering_debug_[Timestamp].txt`) directly to the user's Desktop for error tracking.

## 🛠️ How to Use It

1. Open your Excel workbook and press `ALT + F11` to launch the VBA Developer Window.
2. Click **Insert > Module** from the top menu strip.
3. Copy the code from `ImageBot.vba` in this repository and paste it into the blank module canvas.
4. Close the VBA window to return to your grid layout sheet.
5. Press `ALT + F8`, highlight **`InteractiveMasterImageBotWithLogs`**, and click **Run**.
6. Follow the pop-up prompts to customize your row range, column letter, and grid dimensions!

## 📝 What Solved the "Even/Odd Row Skipping" Issue?
Excel natively processes `Picture in Cell` elements via background graphics threads. When traditional macros run through loops sequentially, Excel's layout engine frequently drops command frames on alternate (even) rows while updating the spreadsheet index. 

This utility addresses that platform constraint by pairing an expanded shape object verification check with an autonomous clipboard-paste backup mechanism. If a cell fails to release its image via standard system commands, the script seamlessly deploys the bypass, clears the background cache, and records the event execution logs directly to your Desktop.

## 📄 License
This project is open-source and available under the [MIT License](LICENSE).
