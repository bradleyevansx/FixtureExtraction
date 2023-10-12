using System.Collections.ObjectModel;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

public class WebScrape
{
    public async Task ExecuteAsync()
    {
        var filePath = new FileInfo("C:\\Users\\bradl\\POL\\WebScraping\\hubbardton.xlsx");
        
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        
        ExcelPackage package = new ExcelPackage(filePath);
        
        var worksheet = package.Workbook.Worksheets["Sheet1"];

        var chromeOptions = new ChromeOptions();
        chromeOptions.AddArguments("headless");


        var allProductLinks = new List<string>();

        using (var driver = new ChromeDriver(chromeOptions))
        {
            for (int i = 1; i <= 149; i++)
            {
                var currCollectionUrl = (string)worksheet.Cells[1, i].Value;
                driver.Navigate().GoToUrl(currCollectionUrl);
                await Task.Delay(3000);
                
                ReadOnlyCollection<IWebElement> productLinks = driver.FindElements(By.CssSelector("a.productItem-images-K2A.bg-white.relative.px-4.pt-3"));
                
                foreach (var link in productLinks)
                {
                    string href = link.GetAttribute("href");
                    allProductLinks.Add(href);
                }

            }

            var row = 2;
            foreach (var link in allProductLinks)
            {
                worksheet.Cells[row, 1].Value = link;
                row++;
            }
        }
        
        package.Save();
        package.Dispose();
    }

    private async Task<List<string>> FindCollections(ChromeDriver driver)
    {
        var collections = new List<string>();
        
        var findCollections = new[]
        {
            "https://hubbardtonforge.com/collections?page=1&pageSize=45&filterId=1",
            "https://hubbardtonforge.com/collections?page=2&pageSize=45&filterId=1",
            "https://hubbardtonforge.com/collections?page=1&pageSize=45&filterId=2",
            "https://hubbardtonforge.com/collections?page=1&pageSize=45&filterId=3",
            "https://hubbardtonforge.com/collections?page=1&pageSize=45&filterId=4",
            "https://hubbardtonforge.com/collections?page=1&pageSize=45&filterId=5"
        };
        
        foreach (var url in findCollections)
        {
            driver.Navigate().GoToUrl(url);
            await Task.Delay(5000);

            ReadOnlyCollection<IWebElement> readMoreLinks = driver.FindElements(By.XPath("//a[@class='underline mx-auto mb-8 text-bronze']"));
        
            foreach (var link in readMoreLinks)
            {
                string href = link.GetAttribute("href");
                collections.Add(href);
            }
        }

        return collections;
    }
}

class Program
{
    static async Task Main()
    {
        var webScrape = new WebScrape();
        await webScrape.ExecuteAsync();
    }
}