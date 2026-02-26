# Voucher Automation System

A robust desktop application for Nagarkot Forwarders to automate the processing of reimbursement and job-related vouchers.

## Tech Stack
- **Python 3.10**
- **Tkinter** (GUI)
- **Pandas** (Data Transformation)
- **Openpyxl** (Excel Generation)

---

## Installation & Setup (For Developers)

⚠️ **IMPORTANT:** You must use a virtual environment.

1. **Create virtual environment**
   ```bash
   python -m venv venv
   ```

2. **Activate (REQUIRED)**
   - Windows: `venv\Scripts\activate`
   - Mac/Linux: `source venv/bin/activate`

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Run application**
   ```bash
   python reimbursement_app.py
   ```

---

## Building the Executable

This project uses **PyInstaller** to create a standalone `.exe`.

1. **Initialize Environment**: Ensure your `venv` is active and `pyinstaller` is installed.
2. **Build via Spec File**:
   ```bash
   pyinstaller VoucherProcessor.spec
   ```
3. **Locate Executable**: The application will be generated in the `dist/` folder.

---

## Usage
Refer to the [USER_GUIDE.md](USER_GUIDE.md) for detailed instructions on how to use the application features.

---

## Notes
- **Local History**: The app tracks processed Transaction IDs in `download_history.json`.
- **Assets**: The logo and expense code references are bundled within the executable.
