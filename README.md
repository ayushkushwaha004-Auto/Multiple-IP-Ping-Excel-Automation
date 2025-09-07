# 📊 Multiple-IP Ping Automation with PowerShell + Excel  

This PowerShell script automates network health checks by reading a list of IPs from Excel, pinging them multiple times, and saving the results with color-coded status.  

---

## 🚀 Features
- Reads IPs from Excel (Column A).  
- Pings each IP **10 times** for reliability.  
- Writes results back to Excel with:  
  - ✅ Green = Up  
  - ❌ Red = Down  
- Saves results into a **timestamped Excel file** (keeps history).  
- Generates log files for troubleshooting.  
- Can be scheduled via **Windows Task Scheduler** for automation.  

---

## 📂 Files
- `Multiple-IP-Ping-Excel.ps1` → The PowerShell script  
- `All-IP-Source-Sheet.xlsx` → Example Excel input file (Column A = IPs)  
- `Logs/` → Auto-generated logs on each run  

---

## ⚡ Usage
1. Clone/download this repository  
2. Place your IPs in **Column A** of `All-IP-Source-Sheet.xlsx`  
3. Run the script:  
   ```powershell
   .\Multiple-IP-Ping-Excel.ps1
---

## Results will be saved as:

Ping-Result_YYYY-MM-DD_HH-mm-ss.xlsx