# Como executar a automação do Diário Oficial com Selenium

Este projeto tem como objetivo automatizar a extração de marcadores do Diário Oficial utilizando o Selenium WebDriver e o pacote OfficeOpenXml para manipulação de planilhas Excel. Neste guia, vamos apresentar um passo a passo para executar a automação.
## Pré-requisitos
Antes de executar a automação, é necessário ter instalado em sua máquina os seguintes softwares:

- .NET Core SDK
- Google Chrome
- ChromeDriver
- Instalação

Navegue até a pasta do projeto e abra o arquivo Robo.csproj.

No arquivo Robo.csproj, verifique se as referências estão instaladas corretamente:

<PackageReference Include="Selenium.WebDriver" Version="3.141.0" />
<PackageReference Include="Selenium.WebDriver.ChromeDriver" Version="95.0.4638.54" />
<PackageReference Include="EPPlus" Version="5.6.4" />
<PackageReference Include="Syncfusion.Pdf.NET" Version="19.3.0.57" />

# Execução

Abra o prompt de comando e navegue até a pasta do projeto.

Execute o comando "dotnet run" para executar a automação.

Aguarde até que a automação seja concluída. O resultado da extração de marcadores será salvo em uma planilha Excel, no diretório C:\Users<seu usuário>\Documents\C#\Selenium\Docs\Marcadores_de_Titulo.

# Conclusão

Pronto! Agora você pode executar a automação do Diário Oficial em sua máquina local. Se encontrar algum problema, verifique os pré-requisitos e a instalação.


------------





# How to Run the Official Gazette Automation with Selenium

This project aims to automate the extraction of markers from the Official Gazette using Selenium WebDriver and the OfficeOpenXml package for manipulating Excel spreadsheets. In this guide, we will provide a step-by-step guide to running the automation.

# Prerequisites
Before running the automation, you must have the following software installed on your machine:

- .NET Core SDK
- Google Chrome
- ChromeDriver
- Installation

Navigate to the project folder and open the Robo.csproj file.
In the Robo.csproj file, make sure the following references are installed correctly:
<PackageReference Include="Selenium.WebDriver" Version="3.141.0" />
<PackageReference Include="Selenium.WebDriver.ChromeDriver" Version="95.0.4638.54" />
<PackageReference Include="EPPlus" Version="5.6.4" />
<PackageReference Include="Syncfusion.Pdf.NET" Version="19.3.0.57" />
# Execution

Open the command prompt and navigate to the project folder.
Run the command "dotnet run" to execute the automation.
Wait until the automation is completed. The result of the marker extraction will be saved in an Excel spreadsheet in the directory C:\Users<your user>\Documents\C#\Selenium\Docs\Marcadores_de_Titulo.
# Conclusion
Now you can run the Official Gazette automation on your local machine. If you encounter any problems, check the prerequisites and installation.
