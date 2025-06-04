### 🧠 What happens when you combine Event Viewer analysis with PowerShell and HTML?
A interactive session log report.

---

### 📂 How It Works

This project uses a set of PowerShell scripts to parse relevant event logs and generate `.html` reports. These reports are then opened via a small `VBScript` launcher for ease of use.

---

### ⚙️ Setup Instructions

1. **Clone or download this repository.**
2. **Adjust the output path** for the generated `.html` report in each `.ps1` file.
   You’ll find the path defined near the end of each script — modify it as needed.
3. **(Recommended)** Create a shortcut to `iniciar_portal.vbs` and place it on the Desktop or Start Menu.
   This makes it easy to launch the portal with a double-click.
4. **Run the scripts with appropriate privileges.**
   Some Event Viewer data requires administrative access.

---
