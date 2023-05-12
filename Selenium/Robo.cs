using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Text;
using Syncfusion.Pdf.Parsing;
using Syncfusion.Pdf.Interactive;


namespace Selenium
{
    internal class Robo
    {
        private IWebDriver driver;

        
        string downloadpath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)+ @"\C#\Selenium\Docs\DiarioOficial";

        public Robo()
        {

            InitializeDriver();
        }
        public void InitializeDriver()
        {
            ChromeOptions options = new ChromeOptions();

            options.AddUserProfilePreference("download.default_directory", downloadpath);
            options.AddUserProfilePreference("download.prompt_for_download", false);
            options.AddUserProfilePreference("download.directory_upgrade", true);
            options.AddUserProfilePreference("safebrowsing.enabled", true);

            Directory.CreateDirectory(downloadpath);
            driver = new ChromeDriver(options);

            Log("Inicializando: Robo");
        }
        private void Log(string message)
        {
            string newFolder = "LogRobo";
            string caminho = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "C#", "Selenium", "Docs", newFolder);

            Directory.CreateDirectory(caminho);

            string arquivo = Path.Combine(caminho, "log.txt");
            using (StreamWriter writer = new StreamWriter(arquivo, true, Encoding.Default))
            {
                writer.WriteLine($"{DateTime.Now} >>> {message}");
            }


        }

        public void AccessPage()
        {
            driver.Navigate().GoToUrl("https://doweb.rio.rj.gov.br/");
            Log("Acessando Pagina");

        }

        public void DownloadDiario()
        {
            Log("Tentando realizar o download");
            try
            {
                driver.FindElement(By.XPath("//*[@id=\"imagemCapa\"]//parent::a")).Click();
                Thread.Sleep(10000);
                Log("Download Concluido com sucesso!");
            }
            catch (Exception ex)
            {

                Log($"Dowload falhou: {ex.Message}");
            }
            finally
            {
                driver.Close();
            }
            
            
        }


        public void ExtractPDF()
        {
            var pdfFiles = Directory.GetFiles(downloadpath, "*.pdf").OrderBy(f => new FileInfo(f).LastWriteTimeUtc);


            
            if (!pdfFiles.Any())
            {
                Log("Nenhum arquivo PDF encontrado no diretório.");
                return;
            }

            var latestPdfFile = pdfFiles.Last();

            var document = new PdfLoadedDocument(latestPdfFile);
            Log("Extraindo marcadores do PDF.");
            var bookmarks = document.Bookmarks;

            ExportTitleToExcel(bookmarks);
        }



        public void ExportTitleToExcel(PdfBookmarkBase bookmarks)
        {
            Log("Iniciando a exportação dos marcadores");

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string newFolder = "Marcadores_de_Titulo";
            string caminho = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "C#", "Selenium","Docs", newFolder);





            Directory.CreateDirectory(caminho);
           

            string formattedDate = DateTime.Today.ToString("dd-MM-yyyy");
            string fileName = $"Diario_{formattedDate}.xlsx";
            string filePath = Path.Combine(caminho, fileName);

            using (var excelPackage = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = null;

                //verifica se existe o worksheet ativo, e deleta caso exista
                foreach (var existingWorksheet in excelPackage.Workbook.Worksheets)
                {
                    if (existingWorksheet.Name == "Diario Oficial")
                    {
                        worksheet = existingWorksheet;
                        excelPackage.Workbook.Worksheets.Delete(worksheet);
                        break;
                    }
                }

                worksheet = excelPackage.Workbook.Worksheets.Add("Diario Oficial");

                worksheet.Cells[1, 1].Value = "Titulo Principal";
                worksheet.Cells[1, 2].Value = "Página";
                worksheet.Cells[1, 3].Value = "Data de Execução";

                int currentRow = 2;

                Log("Iniciando a inserção de conteúdo");
                WriteBookmarksToWorksheet(bookmarks, worksheet, ref currentRow);
                Log("Finalizando a inserção de conteúdo");

                excelPackage.SaveAs(new FileInfo(filePath));
            }

            Log("Planilha salva com sucesso");
        }

        private  void WriteBookmarksToWorksheet(PdfBookmarkBase bookmarks, ExcelWorksheet worksheet, ref int currentRow)
        {
           
            if (bookmarks.Count > 0 ) {
                foreach (PdfBookmark bookmark in bookmarks)
                {
                    // Escreve o nome do marcador e o número da página na planilha
                    worksheet.Cells[currentRow, 1].Value = bookmark.Title;
                    worksheet.Cells[currentRow, 2].Value = bookmark.Destination.PageIndex + 1;
                    worksheet.Cells[currentRow, 3].Value = DateTime.Now.ToString();

                   
                    currentRow++;

                    // Chama recursivamente esta função para iterar sobre os sub-marcadores, mesmo que o bookmark atual não tenha sub-marcadores
                    WriteBookmarksToWorksheet(bookmark, worksheet, ref currentRow);
                }
            }

           
        }
        public void Finalizar()
        {
            Task.WaitAll(Task.Delay(1000));
            Log("Finalizando o Robo");
            Console.WriteLine("pressione qualquer tecla para finalizar");
            Environment.Exit(0);
            Console.ReadKey();
        }
    }
        

}




