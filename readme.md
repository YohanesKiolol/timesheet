# Timesheet 🔥

## Installation

### Prerequisites

- **Python**: Ensure you have Python3 installed on your system. You can download it from [python.org](https://www.python.org/downloads/) or use a package manager appropriate for your operating system.

### Steps

1. **Install python-tk with brew**

   ```bash
   brew install python-tk
   ```

2. **Setup environtment**

   - Copy the .env-copy file and rename it to .env. Modify the contents of the .env file to set the desired path and name.
   - Create a virtual environment:
     ```bash
     python3 -m venv .venv
     ```
   - Activate the virtual environment:
     ```bash
     source .venv/bin/activate
     ```
   - Install the required dependencies:
     ```bash
     python3 -m pip install -r requirements.txt
     ```

3. **Run App**

   ```bash
   python3 main.py
   ```

4. **Create a .exe (Optional)**
   - Build the app
     ```bash
        pyinstaller --onefile main.py
     ```
   - Run the main.exe inside /dist folder
