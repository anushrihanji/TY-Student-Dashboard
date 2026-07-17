# TY Student Dashboard

## Setup & Run

1. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

2. Place Excel files in the `data/` folder:
   - `TY-A_Attendances.xlsx`
   - `TYB_Attendance.xlsx`
   - `TY-C_Attendance.xlsx`
   - `TY-A_Marks.xlsx`
   - `TY-B_Marks.xlsx`
   - `TY-C_Marks.xlsx`

3. Run the app:
   ```
   python app.py
   ```

4. Open http://localhost:5000 in your browser.

## Passwords
- Division A: `divA@2024`
- Division B: `divB#2024`
- Division C: `divC$2024`

## Notes
- Any changes to Excel files are reflected immediately on refresh (no restart needed).
- Attendance % is calculated from lecture counts directly; changes to Excel update automatically.
- ISE-1 = CO1 + CO2 (out of 40), ISE-2 = CO3 + CO4 (out of 40), Average = (ISE1+ISE2)/2.
- AB in both CO1+CO2 → ISE = AB. If ISE1=AB in Average calc, treat as 0.
