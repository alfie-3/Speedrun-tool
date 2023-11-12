using Microsoft.Office.Interop.Excel;
using NicoSpeedrunTool;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using static System.Net.WebRequestMethods;
using Excel = Microsoft.Office.Interop.Excel;

public class SpeenrunGetter {
    //Excel Variables
    static Excel.Application oXL;
    static Excel._Workbook oWB;
    static Excel._Worksheet oSheet;
    static string docName = "Active Player Satistics";

    static int writeColumn = 0;
    static Range bottomRight;

    //Gathered item variables
    static readonly string top50URL = "https://www.speedrun.com/games?page=1&platform=&sort=mostactive";
    static List<GameInfo> games;
    static int amountOfGamesToCollect = 200;

    static void Main() {
        Tools.WriteMessage($"Speedrun Collection Tool V2.1", ConsoleColor.Magenta);

        StartExcel();

        //Starts stopwatch to check how long the process takes
        Stopwatch stopwatch = new Stopwatch();
        stopwatch.Start();

        //Gathers top 50 games from the website using Selenium
        Tools.WriteMessage($"Collecting {amountOfGamesToCollect} games from https://www.speedrun.com/games...", ConsoleColor.Red);
        games = GatherTopGames();

        Tools.WriteMessage($"{games.Count} found, writing to Excel...", ConsoleColor.Red);

        WriteGameData();

        stopwatch.Stop();
        TimeSpan t = TimeSpan.FromMilliseconds(stopwatch.ElapsedMilliseconds);

        Tools.WriteMessage($"Collected games in {t:mm': 'ss'. 'ff}", ConsoleColor.Red);

        ExcelQuit();
    }

    private static void StartExcel() {
        //Gets document path
        var currentDocPath = Path.Combine(Directory.GetCurrentDirectory(), $@"..\Output/{docName}.xlsx");
        CreateOutputFileIfNotAlreadyCreated();

        oXL = new Excel.Application {
            //Visible = true,
            UserControl = false,
            DisplayAlerts = false
        };

        //Opens if document exists, create new document if it does not
        try {
            oWB = oXL.Workbooks.Open(currentDocPath);
        }
        catch {
            oWB = oXL.Workbooks.Add(Missing.Value);
            oWB.SaveAs2(currentDocPath);
        }

        oSheet = (Excel._Worksheet)oWB.ActiveSheet;

        if (oSheet.Cells[1, 1].Value2 == null)
            oSheet.Cells[1, 1] = "Game";

        //Add current date
        FillInDate();
    }

    private static List<GameInfo> GatherTopGames() {
        var options = new FirefoxOptions();
        options.AddArguments("--headless");

        var service = FirefoxDriverService.CreateDefaultService();
        service.HideCommandPromptWindow = true;

        var fireFoxDriver = new FirefoxDriver(service, options);
        fireFoxDriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(120);

        fireFoxDriver.Navigate().GoToUrl(top50URL);
        Assert.AreEqual("Games - Speedrun", fireFoxDriver.Title);

        //fireFoxDriver.FindElement(By.CssSelector("#qc-cmp2-ui > div.qc-cmp2-footer.qc-cmp2-footer-overlay.qc-cmp2-footer-scrolled > div > button.css-47sehv")).Click();

        int pagesToLoad = (int)Math.Ceiling((double)amountOfGamesToCollect / 90);

        List<GameInfo> games = new List<GameInfo>();

        for (int p = 1; p < pagesToLoad + 1; p++) {

            if (p != 1)
                fireFoxDriver.Navigate().GoToUrl($"https://www.speedrun.com/games?page={p}&platform=&sort=mostactive");

            var gameContainerElements = fireFoxDriver.FindElements(By.XPath("//*[@id=\"app-main\"]/div[4]/div/div[1]/div/div[2]/div[1]/*"));

            for (int i = 0; i < gameContainerElements.Count(); i++) {
                string gameName = gameContainerElements[i].FindElement(By.XPath(".//div/div[1]/a")).GetAttribute("innerHTML");
                string activePlayers = gameContainerElements[i].FindElement(By.XPath(".//div/div[2]")).GetAttribute("innerHTML");

                GameInfo game = new GameInfo {
                    name = gameName,
                    activePlayers = new String(activePlayers.Where(Char.IsDigit).ToArray())
                };

                games.Add(game);
            }
        }

        games = games.GetRange(0, amountOfGamesToCollect);
        fireFoxDriver.Quit();
        return games;
    }

    public static void WriteGameData() {
        int cell = 2;

        while (oSheet.Cells[cell, 1].Value2 != null) {
            if (!CheckIfGameExists(cell)) {
                oSheet.Cells[cell, writeColumn].Value2 = "";
            }

            bottomRight = oSheet.Cells[cell, writeColumn];

            cell++;
        }

        for (int i = 0; i < games.Count; i++) {
            oSheet.Cells[cell, 1].Value2 = games[i].name;
            oSheet.Cells[cell, writeColumn].Value2 = games[i].activePlayers;

            bottomRight = oSheet.Cells[cell, writeColumn];

            cell++;
        }
    }

    public static void FillInDate() {
        DateTime date = DateTime.Today;

        int i = 2;

        while (oSheet.Cells[1, i].Value2 != null) {
            i++;
        }

        writeColumn = i;

        oSheet.Cells[1, i] = date.ToString("dd/MM/yyyy");
    }

    public static void CreateTable() {
        // define points for selecting a range
        // point 1 is the top, leftmost cell
        Excel.Range oRng1 = oSheet.Range["A1"];

        // define the actual range we want to select
        var oRng = oSheet.Range[oRng1, bottomRight];
        oRng.Select(); // and select it

        // add the range to a formatted table
        oRng.Worksheet.ListObjects.AddEx(
            SourceType: Excel.XlListObjectSourceType.xlSrcRange,
            Source: oRng,
            XlListObjectHasHeaders: Excel.XlYesNoGuess.xlYes);
    }

    public static void ResizeTable(ListObject table) {
        // define points for selecting a range
        // point 1 is the top, leftmost cell
        Excel.Range oRng1 = oSheet.Range["A1"];

        // define the actual range we want to select
        var oRng = oSheet.Range[oRng1, bottomRight];

        table.Resize(oRng);
    }

    public static void ExcelQuit() {
        var currentDocPath = Path.Combine(Directory.GetCurrentDirectory(), $@"..\Output/{docName}.xlsx");

        oSheet.Columns.AutoFit();

        if (oSheet.ListObjects.Count == 0)
            CreateTable();
        else {
            var table = oSheet.ListObjects[1];
            ResizeTable(table);
        }

        oXL.UserControl = true;
        oWB.SaveAs2(currentDocPath, CreateBackup: true);
        oWB.Close();

        if (oSheet != null) {
            Marshal.FinalReleaseComObject(oSheet);
            oSheet = null;
        }
        if (oWB != null) {
            Marshal.FinalReleaseComObject(oWB);
            oWB = null;
        }
        if (oXL != null) {
            oXL.Quit();
            Marshal.FinalReleaseComObject(oXL);
            oXL = null;
        }
    }

    public static bool CheckIfGameExists(int cell) {
        foreach (GameInfo game in games.ToList()) {
            if (game.name == oSheet.Cells[cell, 1].Value2.ToString()) {
                oSheet.Cells[cell, writeColumn].Value2 = game.activePlayers;
                games.Remove(game);
                return true;
            }
        }

        return false;
    }

    public static void CreateOutputFileIfNotAlreadyCreated() {

        var path = Path.Combine(Directory.GetCurrentDirectory(), $@"..\Output");

        char tick = '✓';

        if (Directory.Exists(path)) {
            return;
        }
        else {
            Directory.CreateDirectory(path);
            Tools.WriteMessage($"{tick} Output file created at {path}", ConsoleColor.Green);
        }
    }
}
