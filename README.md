# Point Auto Clicker

Point Auto Clicker (PointAC) is a lightweight and modern auto clicker built with **WPF .NET 9** using Microsoft’s **Fluent UI for WPF**.  
It provides a simple and intuitive interface for automating mouse clicks, visualizing click points, and customizing click behavior.

This project was made just for fun over the course of **five days**.  
I’m not a professional .NET or C# developer, so expect a few rough edges or errors here and there.  
The goal was mainly to experiment with modern WPF design, Direct2D rendering, and low-level system interaction.

---

## Features

- **Modern Fluent UI**
  - Uses Microsoft’s Fluent UI framework for WPF to deliver a clean Windows 11-style interface with theme support.
- **Point Visualization**
  - Renders visual click points using **Direct2D**, showing exactly where clicks will occur on the screen.
- **Global Settings**
  - Configure a global click duration.  
  - Choose between *Single* or *Double* click.  
  - Select the mouse button: *Left*, *Right*, *Middle*, or *System Default*.
- **Per-Point Customization**
  - Once points are added, you can adjust each one individually through the points list.
- **Looping**
  - Supports single-run or continuous looping of your click sequence.
- **Normal Auto Clicker Mode**
  - Can function as a simple auto clicker when no points are defined.

---

## How It Works

PointAC uses **P/Invoke** to call low-level Windows APIs (like `user32.dll`) for simulating mouse input.  
This approach enables precise automation, allowing clicks and cursor movements to be generated directly through the system.

Because it interacts with the system at a low level and references **System.Diagnostics** (only for version checking and opening web links),  
some antivirus software may falsely flag it as a Trojan. This is a **false positive** — the app is completely safe to use.

---

## Requirements

- .NET 9 Runtime or later  
- Windows 11 (I don't know if it works in Windows 10 or not; you could figure that out by yourself, it should tho since .NET 9 is supported on Windows 10).

---

## Getting Started

1. Clone or download the repository.  
2. Build the project using Visual Studio or the .NET CLI.  
3. Run the executable to start Point Auto Clicker.  
4. Add and configure your click points.  
5. Start the clicker to automate your sequence.

---

## Technical Overview

| Component | Purpose |
|------------|----------|
| **WPF (.NET 9)** | Core UI framework |
| **Fluent UI for WPF** | Modern Microsoft design and theme support |
| **Direct2D** | Used to render visual point overlays |
| **P/Invoke (user32.dll)** | Handles low-level mouse simulation |

---

## Developer Notes

This project was built out of curiosity to explore:
- The integration of Fluent UI with WPF  
- Direct2D rendering for simple 2D visualization  
- Low-level input handling using P/Invoke  

It’s not meant to be a professional-grade tool — just a fun side project that happens to work pretty well.

---

## License

This project is licensed under the **MIT License**.  
See the [LICENSE](./LICENSE) file for more details.
