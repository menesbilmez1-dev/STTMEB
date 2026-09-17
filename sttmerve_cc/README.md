# STTPROJECT - Excel Komuta Merkezi (Excel Command Center)

A robust bridge application designed to synchronize real-time data from remote clients directly into organized Excel spreadsheets.

## 🚀 Overview
**STTPROJECT** acts as a high-speed HTTP listener that waits for command payloads from any compatible client. Once a command is received, the application identifies the specific Excel sheet targeted, finds the intersection of the **Category** (Main Header) and the **Type** (Subheader), and injects the data into the next available row.

## ✨ Key Features
- **Real-time HTTP Listener**: Runs a lightweight local server (default port 8080) to receive JSON payloads.
- **Automated Excel Management**: 
    - Targeted sheet selection.
    - Intelligent Header Mapping: Uses **Category** as the Main Header and **Type** as the Subheader to locate columns.
    - Dynamic row creation to prevent data overwriting.
- **Live Logging System**: A real-time console within the UI to monitor server status, received requests, and Excel operations.
- **Flexible Configuration**: Easily define your Excel structure and command categories via a configuration file.
- **WPF Desktop Interface**: A clean, intuitive management dashboard for monitoring and manual control.

## 🛠 Tech Stack
- **Backend/Core**: C# / .NET (WPF)
- **Architecture**: MVVM (Model-View-ViewModel)
- **Communication**: HTTP/REST (JSON)
- **Data Storage**: Local Repository & Excel (`.xlsx`)
- **Compatibility**: Works with any device/client capable of sending HTTP POST requests.

## 🔄 Workflow
1. **Configuration**: Define your Excel structure and command categories in the config file.
2. **Startup**: Launch the desktop application and select your target Excel file.
3. **Server Activation**: Start the HTTP server within the app.
4. **Data Transmission**: A client sends a `POST` request with a JSON payload:
   ```json
   {
     "sheetName": "SensorData",
     "category": "Temperature",
     "type": "Sensor_A",
     "value": "24.5"
   }
   ```
5. **Automation**: The app receives the data, opens the `"SensorData"` sheet, finds the column where the **Category** is `"Temperature"` and the **Type** is `"Sensor_A"`, and writes `"24.5"` to the next available row.

## 📂 Project Structure
- `Services/`: Core logic (Excel, HTTP Server, Config, Settings).
- `ViewModels/`: UI logic and data binding.
- `Models/`: Data structures and command payloads.
- `Helpers/`: Converters and utility classes.
- `Views/`: WPF XAML interface.
