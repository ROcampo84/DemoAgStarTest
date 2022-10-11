# DemoAgStarTest
Source code for demonstration purposes


## Requirements
 - VS Community Edition: https://visualstudio.microsoft.com/vs/community/
 - Node.js: https://nodejs.org/en/download/
 - Appium: cmd ==> npm install -g appium
 - WinAppDriver: https://github.com/Microsoft/WinAppDriver 
 - AgStar .NET version.


 ### Steps for configuration
    - AgStar .NET version installed locally (or another desktop application with its controls identified).
    - Run WAD (Windows Application Driver) must be listening for requests at http://127.0.0.1:4723/
    - Switch to Developer mode in Windows 10 (Windows 10 is a requirement to run WinAppDriver)
    - On Visual Studio project:
      - You must set the application target path (where executable is located).
      - You must set the Excel path (with test values).
