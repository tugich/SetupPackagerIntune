# SetupPackagerIntune



<!-- ABOUT THE PROJECT -->
## About The Project

 This application helps you to package your setup files for Intune faster and easier with the official Win32 Content Prep Tool, without entering any commands manually in the console:

![alt text](https://raw.githubusercontent.com/tugich/SetupPackagerIntune/main/Featured_Image.png)



<!-- GETTING STARTED -->
## Getting Started

### Prerequisites

Download Win32 Content Prep Tool at [https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool](https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool).

### Installation

1. Download Win32 Content Prep Tool (see Prerequisites)
2. Copy **IntuneWinAppUtil.exe** in the same folder like SetupPackager.exe - Don't change the name of the wrapper



<!-- USAGE EXAMPLES -->
## Usage

**How to video:**<br>
https://blog.tugi.ch/scripts-and-tools/setup-packager-for-intune

<br>

1. Create the folder structure - My recommandation:

- Software Packages
- **Software Packages / 7-zip** (main folder)
- **Software Packages / 7-zip** / Installer (which contains the installer of the app, *.exe or *.msi)
- **Software Packages / 7-zip** / Package (for the package file / *.intunewin)
- **Software Packages / 7-zip** / Documentations (for documentations etc.)

2. Download prerequisites

3. Run the app `SetupPackager.exe` and drop your main folder to the upload icon.

4. Select the installer file in the dropdown - EXE, MSI or the installer script

5. Click on `Create Intune package`



<!-- LICENSE -->
## License

Distributed under the MIT License. See `LICENSE.txt` for more information.



<!-- CONTACT -->
## Contact

TUGI - [contact@tugi.ch](mailto:contact@tugi.ch)<br>
Project Link: [https://blog.tugi.ch/scripts-and-tools/setup-packager-for-intune](https://blog.tugi.ch/scripts-and-tools/setup-packager-for-intune)

<p align="right">(<a href="#readme-top">back to top</a>)</p>
