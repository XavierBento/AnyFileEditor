# AnyFile Editor (TxtOrganizer)

🚀 **AnyFile Editor** (formerly *TxtOrganizer*) is a modern, lightweight, and open-source multi-format text editor built with **.NET 8.0 (WPF, MahApps.Metro, ControlzEx)**.  
It is designed for writers, students, and developers who need a **fast, tabbed, theme-aware editor** that supports **RTF, TXT, DOCX, ODT** and more.

---

## ✨ Features

- 📝 **Multi-format support**:  
  Open, edit, and save documents in:
  - `.txt`
  - `.rtf`
  - `.docx`
  - `.odt`

- 📂 **Tabbed editing**  
  Work with multiple files in a tabbed interface (like modern browsers).

- 🎨 **Themes & Dark Mode**  
  Full light/dark theme switching powered by MahApps.Metro.

- 🎨 **Custom Color Palette**  
  Pre-defined colors + color picker for rich text formatting.

- 🔠 **Text formatting tools**  
  Bold, Italic, Underline, Alignments, Line spacing, and more.

- 🖼️ **Embedded icons**  
  Toolbar icons embedded via `pack://application` URIs (no missing resource issues).

- 📑 **Split View mode**  
  Edit two files side-by-side in perfectly aligned editors.

- 🖨️ **Printing support**  
  Print documents with A4 preset formatting.

---

## 📸 Screenshots

> *(Add screenshots of the UI here – Dark/Light mode, tabbed view, split view, etc.)*

---

## 🏗️ Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/YOUR-USERNAME/AnyFileEditor.git
   ```
2. Open the solution in **Visual Studio 2022** (with .NET 8.0 SDK installed).
3. Restore NuGet packages.
4. Build & run.

---

## 📦 Dependencies

- [.NET 8.0 SDK](https://dotnet.microsoft.com/en-us/download)
- [MahApps.Metro](https://github.com/MahApps/MahApps.Metro)
- [ControlzEx](https://github.com/ControlzEx/ControlzEx)
- [DocumentFormat.OpenXml](https://github.com/OfficeDev/Open-XML-SDK)
- [OpenXmlPowerTools](https://github.com/OfficeDev/Open-Xml-PowerTools)

---

## 🛠️ Project Structure

```
AnyFileEditor/
│
├── Assets/               # App icons and toolbar icons
├── Models/               # Core models (DocTab, etc.)
├── Views/                # XAML files (MainWindow.xaml, etc.)
├── MainWindow.xaml.cs    # Entry point & event handling
├── MainWindow.Tabs.cs    # Tab management logic
├── MainWindow.Files.cs   # File open/save handlers
├── MainWindow.UI.cs      # UI helpers (themes, colors, etc.)
├── AnyFileEditor.csproj  # Project file
└── README.md             # This file
```

---

## 🧪 Tests

Unit tests for file operations and tab persistence are included under the `Tests/` folder.

Run tests with:
```bash
dotnet test
```

---

## 🤝 Contributing

Contributions are welcome!  
To contribute:

1. Fork this repo
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

---

## 📜 License

This project is licensed under the **MIT License** – see the [LICENSE](LICENSE) file for details.

---

## 💡 Acknowledgements

- Thanks to the [MahApps.Metro](https://mahapps.com) and [ControlzEx](https://github.com/ControlzEx/ControlzEx) teams.
- Inspired by the need for a **free, portable, and theme-aware** text editor.
