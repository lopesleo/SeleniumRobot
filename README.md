How to Run the Official Gazette Automation with Selenium

This is a project for automating the extraction of markers from the Official Gazette using Selenium WebDriver and the OfficeOpenXml package for manipulating Excel spreadsheets. This guide aims to provide a step-by-step guide to running the automation on another machine.

Prerequisites
Before running the automation, you must have the following software installed on your machine:

.NET Core SDK
Google Chrome
ChromeDriver

Installation
Navigate to the project folder and open the Robo.csproj file. In the Robo.csproj file, make sure the following references are installed correctly:

<PackageReference Include="Selenium.WebDriver" Version="3.141.0" />
<PackageReference Include="Selenium.WebDriver.ChromeDriver" Version="95.0.4638.54" />
<PackageReference Include="EPPlus" Version="5.6.4" />
<PackageReference Include="Syncfusion.Pdf.NET" Version="19.3.0.57" />
Execution
Open the command prompt and navigate to the project folder. Run the command "dotnet run" to execute the automation. Wait until the automation is completed. The result of the marker extraction will be saved in an Excel spreadsheet in the directory C:\Users<your user>\Documents\C#\Selenium\Docs\Marcadores_de_Titulo.

Conclusion
Now you can run the Official Gazette automation on your local machine. If you encounter any problems, check the prerequisites and installation.