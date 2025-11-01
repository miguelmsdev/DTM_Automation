# DTM (Debris Transport Manifest) Auto-Filler
Automates data entry for Brazil’s MTR – Manifesto de Transporte de Resíduos (here called DTM)—the licensed government system that tracks Senders, Transporters, and Receivers of construction waste.

What it does:
Reads authorization sheets (A4) from an Excel list of authorizations (filled with a barcode scanner).
Fills DTM forms automatically in the government portal: receiving date, driver info, vehicle plates—either auto-synced to the authorization date or overridden by the user.
Runs hands-free while you monitor progress.

Why it matters:
Throughput: ~1,000 DTMs/day handled reliably.
Speed: ~1 sec per DTM vs. ~11 sec manually → ~90% faster.
Time saved: turns ~3 hours of repetitive work into ~18 minutes of light supervision.
Operational impact: less overtime, fewer data-entry errors, and reassignable headcount to higher-value tasks.

Key features:
Smart defaults (auto receiving date) with easy manual overrides (date, driver, plates).
Error handling & retries for flaky network/portal responses.
Progress logging so you can audit what was submitted.

Typical workflow:
Scan A4 authorizations → get an Excel file with IDs and metadata.
Launch the program and point it to the Excel file.
Choose whether to auto-use authorization dates or set a custom receiving date; optionally set drivers/plates.
Start and monitor—logs show status and any items needing attention.

Tech stack:
Python for automation & data handling (e.g., pandas, openpyxl).
Automation layer (e.g., Selenium/PyAutoGUI) to drive the portal safely.
.xlsx import with validation and ID de-duplication.

Compliance & safety:
Respects the official DTM/MTR workflow; no scraping of restricted data.
Validates required fields before submission and logs every entry for audit.

Possible Future Improvements:
Headless mode & queue dashboard
Per-submission screenshots for audits
Role-based presets (per transporter/receiver)

Below is a GIF of the program in action:
![GitHubMTRAutomationDTM-1](https://github.com/miguelmsdev/DTM_Automation/assets/83340893/d5c8a6d8-a925-4e57-89b2-ecee417a6699)

The project was written using Python for the coding logic, Selenium for the bridge with the Chrome browser and website, and CustomTKinter for the GUI interface to make an easier and user-friendly way to run the code, plus some small generic libraries.

The GUI interface:
![GIT_MTR_DTM_GUI (4)](https://github.com/miguelmsdev/DTM_Automation/assets/83340893/da7b8dd8-4423-4b25-b1b8-0da58148bc01)

It took around 5 days working part-time to develop the algorithm and then write the code, the same time it would take the code to pay for itself time-wise, and from there on to bring profit for the company and more free time for me to invent new solutions while being available for more brain-needed challenges.

It is not fully-fledged and needs constant maintenance as does every third-party dependant software, the website is constantly changing references and elements' paths, sometimes refactoring almost all of it.
I've paused updating the code after a year and a half of use since contracts have hit the estimated deadlines, but it proved useful saving time and I learned a lot during the process of developing it.
