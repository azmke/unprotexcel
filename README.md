# Unprotexcel

**Unprotexcel** is a lightweight Java-based tool built with Maven designed to remove editing protection from Excel files, perfect for when you have forgotten the password. With a user-friendly interface and robust features, this tool ensures you can access and modify your Excel files with ease.

<div align="center">
    <img src="docs/assets/gui_screenshot.PNG" alt="Tool Interface" width="787">
</div>

## Features

Unprotexcel offers a variety of features to quickly and easily unlock your Excel files:

- **🔍 File Selection:** Choose an Excel file via button, drag-and-drop, or menu.
- **🔓 Unlocking:** Easily unlock protected Excel files with a single click.
- **💾 Save Option:** Save the unlocked file to your desired location.
- **🌗 Themes:** Switch between light and dark themes for a comfortable user experience.
- **🌍 Languages:** Toggle between English and German for a localized experience.

## Download

You can download the latest release of Unprotexcel for your operating system from the links below:

<div align="left">
  <table>
    <tr>
      <th>Windows</th>
      <th>Mac</th>
      <th>Linux</th>
    </tr>
    <tr>
      <td align="center">
        <a href="https://github.com/yourusername/excel-sheet-unlocker/releases/download/latest/windows-installer.exe">
          <img src="https://img.icons8.com/color/48/000000/windows-10.png" alt="Windows"/>
          <br/>
          Download for Windows
        </a>
      </td>
      <td align="center">
        <a href="https://github.com/yourusername/excel-sheet-unlocker/releases/download/latest/mac-installer.dmg">
          <img src="https://img.icons8.com/color/48/000000/mac-logo.png" alt="Mac"/>
          <br/>
          Download for Mac
        </a>
      </td>
      <td align="center">
        <a href="https://github.com/yourusername/excel-sheet-unlocker/releases/download/latest/linux-installer.sh">
          <img src="https://img.icons8.com/color/48/000000/linux.png" alt="Linux"/>
          <br/>
          Download for Linux
        </a>
      </td>
    </tr>
  </table>
</div>

## Usage

Follow these simple steps to use Unprotexcel:

1. **Select File:** Click the "Choose File" button, use drag-and-drop, or go to "File > Open file" to select an Excel file.
2. **Unlock:** Click the "Unlock" button.
3. **Save:** A dialog will prompt you to choose where to save the unlocked file. Select the destination and save.

## Building

### Prerequisites
Ensure you have the following installed:
- [Java Development Kit (JDK)](https://www.oracle.com/java/technologies/javase-jdk11-downloads.html)
- [Apache Maven](https://maven.apache.org/)

### Installation

1. **Clone the repository:**
    ```bash
    git clone https://github.com/azmke/unprotexcel.git
    cd excel-sheet-unlocker
    ```

2. **Build the project with Maven:**
    ```bash
    mvn clean install
    ```

3. **Run the application:**
    ```bash
    java -jar target/unprotexcel-0.1.0-SNAPSHOT.jar
    ```

## Contributing

We welcome contributions from the community! Here's how you can help:

### Reporting Bugs
If you encounter a bug, please open an [issue](https://github.com/azmke/unprotexcel/issues).

### Suggesting Features
We are always open to new ideas. If you have a feature request, please open an [issue](https://github.com/azmke/unprotexcel/issues).

### Code Contributions
1. **Fork the repository** on GitHub.
2. **Create a feature branch:**
    ```bash
    git checkout -b feature/AmazingFeature
    ```
3. **Commit your changes:**
    ```bash
    git commit -m 'Add some AmazingFeature'
    ```
4. **Push to the branch:**
    ```bash
    git push origin feature/AmazingFeature
    ```
5. **Open a Pull Request**.


## License

This project is licensed under the terms of the GNU General Public License v3.0. See the [LICENSE](./LICENSE) file for details.

## Disclaimer

By using this tool, you agree to the terms and conditions outlined in the [DISCLAIMER.md](DISCLAIMER.md).
