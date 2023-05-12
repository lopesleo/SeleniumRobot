using Selenium;

var web = new Robo();
web.AccessPage();
web.DownloadDiario();
web.ExtractPDF();
web.Finalizar();



