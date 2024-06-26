# VFDChrome
A Chrome driver based python script to automate printing for TAX automation software VFDBox. The selenium based chrome driver solution uses the "Save as PDF" google chrome printer driver to save the PDF to tax automation folder for fiscalization of receipts, invoices and credit notes without users intervention upon clicking a print button or `CTRL + P` on any web page.

# Python Libraries
This project was built with python 3.6 using the following libraries:
- `json, os, sys, psutil`
- `subprocess`
- `selenium`
- `chromedriver_autoinstaller_fix`
- `ctypes`
- `threading`
- `datetime`

# Freezing the project for distribution
The project is already frozen. The contents of the cloned repository can be copied to `C:\\VFDChrome`, then make a shortcut of `VFDChrome.exe` to your desktop or pin to task manager. If you wish to freeze the project, clone the repository, isolate `VFDChrome.pyw`, ensure the libraries are installed on your computer otherwise, use `pip install **library`, then freeze the script with your choice python freezer. We have also included a direct installer [kapticVatracSetup.exe](https://github.com/onuh/VFDChrome/raw/main/installer/kapticVatracSetup.exe).

# Running the project
To run the project, you must be connected to internet at first run so the setup can download and install necessary files, after which you can enjoy the solution offline if you run your business on localhost. The solution always connects to the chrome driver update server to get the latest version for your installed google chrome web browser. If you already have google chrome installed, at first launch, it may take a while to come up because it must download chrome driver for your version of installed google chrome. A Google Chrome Web browser installer is also included as part of setup to install chrome should you run it on a computer without google chrome installed. Once lauched successfully, it opens google search page in `kiosk mode`. Close the app using `CTRL + F4`, go to `C:\\VFDChrome` directory and edit the address in `config.ini` to your POS hosting provider, save the file and relaucnch the app. URL of your hosting provider is now loaded. Use google chrome shortcuts to navigate web pages since solution is in `kiosk mode`. `Alt + ->` to navigate foward, `Alt + <-` to navigate backward, `Alt + home` to return to history home.  To use it for automation of Tax, contact the team at [Softrust Technologies Limited](https://softrust.com.ng)