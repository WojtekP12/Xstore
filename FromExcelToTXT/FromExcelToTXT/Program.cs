using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO.Compression;
using System.Threading;

namespace FromExcelToTXT
{
    class Program
    {
        static string PrerequisitiesPath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.Parent.FullName + @"\XStoreInstallation\Prerequisites\";
        static string XStoreInstallationToZipPath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.Parent.FullName + @"\XStoreInstallation";
        static string XstorePackagesFolder = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.Parent.FullName + @"\";
        private static string storeCode;
        private static string storeType = "";
        private static string brand = "";
        private static string numberOfRegisters;
        private static string AJBIP;
        private static string city;
        private static string state;
        private static string country;
        private static int currentRegisterNumber;
        private static int brandSheet;

        const string AMERROLLOUT = "AMER Rollout - Installation Specs_AMER_v57.xlsx";
        const string XSTOREPROPS = "Xstore_properties_configuration_checks4.5+_AMER_v.20.xlsx";
        const string INSTALLATIONCONFIG = "XStoreInstallationConfig.ini";

        static void Main(string[] args)
        {
            Console.WriteLine("Type the store code: ");
            storeCode = Console.ReadLine();
            GetBasicStoreData();

            GetAJBIP();

            for (int i = 1; i <= Convert.ToInt16(numberOfRegisters); i++)
            {

                XStorePropertiesToTxt(i);

                FillInUserProp(i);

                CommaConfigSwitch(i);

                CopyMNTFiles(i);


                USERPROPtoTXT(1);
                USERPROPtoTXT(2);
                USERPROPtoTXT(3);
                AMERtoTXT();

                currentRegisterNumber = i;
                XStorePropertiesToTxt(currentRegisterNumber);

                Thread.Sleep(500);
                DateTime dt = DateTime.Now;
                killExcelProc();
                ZipFile.CreateFromDirectory(XStoreInstallationToZipPath, XstorePackagesFolder + "KA" + storeCode + "R" + String.Format("{0:00}", i) + "_" + String.Format("{0:yyyy.MM.dd}", dt), CompressionLevel.Optimal, true);
                killExcelProc();
            }
        }

        public static void CopyMNTFiles(int register)
        {
            string baseMNTDir = PrerequisitiesPath + @"Base MNT\";

            if (register == 1)
            {
                try
                {
                    System.IO.DirectoryInfo di = new DirectoryInfo(baseMNTDir);

                    Directory.Delete(baseMNTDir,true);
                    Directory.CreateDirectory(baseMNTDir);

                    string sourceFolder = XstorePackagesFolder + @"All MNT\" + brand;
                    string outputFolder = baseMNTDir + brand;
                    Directory.CreateDirectory(outputFolder);
                    Thread.Sleep(3000);
                    new Microsoft.VisualBasic.Devices.Computer().FileSystem.CopyDirectory(sourceFolder, outputFolder, true);

                    foreach (DirectoryInfo directory in di.GetDirectories("OLD", SearchOption.AllDirectories))
                    {
                        foreach (FileInfo file in directory.GetFiles())
                        {
                            file.Delete();
                        }
                        directory.Delete(true);
                    }                
                }
                catch (Exception e)
                {
                    Console.WriteLine("Cos z kopiowaniem mntkow do Base MNT " + e.Message);
                    Console.Read();
                }
            }

        }

        private static void CommaConfigSwitch(int i)
        {
            string fullPath = PrerequisitiesPath + INSTALLATIONCONFIG;
            string commaLine = "CommaInstallation";

            var lineNumber = File.ReadAllLines(fullPath).Select((text, index) => new { text, line = index + 1 }).Where(x => x.text.Contains(commaLine)).FirstOrDefault();

            var loadedTxt = File.ReadAllLines(fullPath);

            if (i == 1)
            {
                loadedTxt[lineNumber.line - 1] = "CommaInstallation = true";
            }
            else
            {
                loadedTxt[lineNumber.line - 1] = "CommaInstallation = false";
            }

            File.WriteAllLines(fullPath, loadedTxt);

        }

        private static string cutAfterEqual(string s)
        {
            int l = s.IndexOf("=");
            if (l > 0)
            {
                return s.Substring(l + 1, s.Length - l - 1);
            }
            return "";
        }

        private static string getValueFromExcelField(Excel.Workbook wb, int sheetNumber, int row, int column)
        {
            System.Globalization.CultureInfo oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            string value;
            Excel.Worksheet excelSheet = (Excel.Worksheet)wb.Sheets[sheetNumber];
            Excel.Range valueRange = (Excel.Range)excelSheet.Cells[row, column];
            

            if(valueRange.Value2!=null)
            {
                value = valueRange.Value2.ToString();
                //killExcelProc();
                return value;
            }
            else
            {
                //killExcelProc();
                return null;
            }
            
        }

        public static void USERPROPtoTXT(int sheet)
        {
            string excelFile = PrerequisitiesPath + "UserProp.xlsx";
            string key;
            string value;
            string line;

            List<string> list = new List<string>();
            int i = 3;

           
            System.Globalization.CultureInfo oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            Excel.Application excel = new Excel.Application();
            Excel.Workbook wb = excel.Workbooks.Open(excelFile);

            key = getValueFromExcelField(wb, sheet, i, 1);
            value = getValueFromExcelField(wb, sheet, i, 2);

            while (key != null)
            {
                if (sheet == 2)
                {
                    line = (key + " " + value);
                }
                else
                {
                    line = (key + value);
                }

                list.Add((line));
                i++;
                key = getValueFromExcelField(wb, sheet, i, 1);
                value = getValueFromExcelField(wb, sheet, i, 2);
            }

            if (sheet == 1)
            {
                list.Add("");
                key = getValueFromExcelField(wb, sheet, 8, 1);
                value = getValueFromExcelField(wb, sheet, 8, 2);
                line = (key + "=" + value);
                list.Add((line));
            }

            string[] lines = list.ToArray();
            System.IO.File.WriteAllLines(PrerequisitiesPath + "USERPropS" + sheet.ToString() + ".txt", lines);
            killExcelProc();
        }

        private static void fillExcelField(Excel.Workbook wb, int sheetNumber, int row, int column, string newValue)
        {    
            Excel.Worksheet excelSheet = (Excel.Worksheet)wb.Sheets[sheetNumber];

            Excel.Range range = (Excel.Range)excelSheet.Cells[row, column];
            range.Value2 = newValue;

            wb.Save();
        }

        public static void FillInUserProp(int registerNumber)
        {          
            string userPropFile = PrerequisitiesPath + "UserProp.xlsx";

            System.Globalization.CultureInfo oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            Excel.Application excel = new Excel.Application();
            Excel.Workbook wb = excel.Workbooks.Open(userPropFile);

            if (registerNumber == 1)
            {
                fillExcelField(wb, 1, 3, 2, "lead");
            }
            else
            {
                fillExcelField(wb, 1, 3, 2, "nonlead");
            }

            fillExcelField(wb, 1, 4, 2, registerNumber.ToString());
            fillExcelField(wb, 1, 5, 2, "KA" + storeCode + "R01");
            fillExcelField(wb, 1, 6, 2, storeCode);
            fillExcelField(wb, 1, 8, 2, AJBIP);
            fillExcelField(wb, 1, 9, 2, registerNumber.ToString());

            if (numberOfRegisters == "1")
            {
                fillExcelField(wb, 1, 9, 2, "KA" + storeCode + "R01");
                fillExcelField(wb, 3, 4, 2, "FALSE");
            }
            else
            {
                fillExcelField(wb, 1, 9, 2, "KA" + storeCode + "R02");
                fillExcelField(wb, 3, 4, 2, "TRUE");
            }
            killExcelProc();
        }

        public static string RemoveSpace(string line)
        {
            string newLine = "";
            for (int i = 0; i < line.Length; i++)
            {
                if (line[i] == '\n')
                {
                    i += 1;
                }
                newLine += line[i];

            }
            return newLine;
        }

        public static void GetAJBIP()
        {
            int i = 2;
            string excelFile = PrerequisitiesPath + "AJB Store terminal Listing.xlsx";

            System.Globalization.CultureInfo oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            Excel.Application excel = new Excel.Application();
            Excel.Workbook wb = excel.Workbooks.Open(excelFile);


            string x = getValueFromExcelField(wb, 2, i, 2);
            

            if (storeType == "SIS")
            {
                AJBIP = "xx.xx.xxx.xxx";
                return;
            }

            try
            {
                while (x != null)
                {
                    if (x == storeCode)
                    {
                        AJBIP = getValueFromExcelField(wb, 2, i, 3);
                        AJBIP = AJBIP.Replace("xx", "133");
                        break;
                    }
                    i++;                  
                    x = getValueFromExcelField(wb, 2, i, 2);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("Couldn't find the AJB Ip for the store provided. Press any key");
                Console.WriteLine("AJB Ip is gonna be set to xx.xx.xxx.xxx.");
                Console.WriteLine("Press any key");
                AJBIP = "xx.xx.xxx.xxx";
                Console.ReadKey();
            }

            if (AJBIP == null)
            {
                Console.WriteLine("NULL - AJB IP not available in the excel file!");

            }

            killExcelProc();
        }






        //CRAP!
        public static void GetActualMNTFolderNames()
        {
            

            string AMERTxtFile = PrerequisitiesPath + "AMER.txt";
            var loadedTxt = File.ReadAllLines(AMERTxtFile);
            var mntFromExcel = "";

            mntFromExcel = cutAfterEqual(loadedTxt[6]);


            // MNT FILES   
            string currentMntFolderName = "";
            string newMntFolders = "";
            string[] stringSeparators = new string[] { ".zip" };
            string tempBrandShortcut = "";
            var mnt = mntFromExcel.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries);

            // 07.03.2016
            string tempMntFile = "";
            string tempMontrealFooter = "";
            foreach (var item in mnt)
            {
                if (item.Contains("CITY"))
                {
                    tempMntFile = item;
                    mnt = mnt.Where(val => val != item).ToArray();
                }

                // Kanada montreal footer bullshit
                if (item.Contains("GG_COUNTRY_CA_MONTREAL_FOOTER"))
                {
                    tempMontrealFooter = item;
                    mnt = mnt.Where(val => val != item).ToArray();
                }
            }
            //

            //tniemy nazwe po podlogach. 
            for (int i = 0; i < mnt.Length; i++)
            {
                string temp = "";
                var names = mnt[i].Split(new string[] { "_" }, StringSplitOptions.RemoveEmptyEntries);
                if (names[0] == "GG")
                {
                    names[0] = "Gucci";
                }
                tempBrandShortcut = names[0];
                for (int j = 0; j < names.Length - 1; j++)
                {


                    //if ((names[j] == "US" || names[j] == "AMER"))  //brand.ToLower() != "slp" && 
                    //{
                    //    continue;
                    //}
                    //else if (names[j] == "AMER")
                    //{
                    //    continue;
                    //}

                    // Z MX i CA?
                    if (names[j] == "AMER")
                    {
                        continue;
                    }

                    if (names[j].ToLower() == "out")
                    {
                        names[j] = "outlet";
                    }
                    temp += names[j] + @"\";
                }

                string[] dirs = new string[0];

                try
                {
                    dirs = Directory.GetDirectories(PrerequisitiesPath + @"Base MNT\" + temp);
                }
                catch (Exception)
                {
                    Console.WriteLine("Prawdopodobnie brak .zip w excelu. Ewentualnie znowu dali OUT zamiast OUTLET");
                    Console.ReadKey();
                    Environment.Exit(0);
                }

                if (dirs.Length == 0)
                {
                    Console.WriteLine("Pewnie niewypakowane mntki " + temp);
                    Console.ReadKey();
                    Environment.Exit(0);
                }

                foreach (string dir in dirs)
                {
                    if ((dirs.Length == 1 && new DirectoryInfo(dir).Name.ToLower() == "old"))
                    {
                        Console.WriteLine("Pewnie niewypakowane mntki " + temp);
                        Console.ReadKey();
                        Environment.Exit(0);
                    }
                    if (new DirectoryInfo(dir).Name.ToLower() != "old")
                    {
                        currentMntFolderName += dir;
                        mnt[i] = Path.GetFileName(currentMntFolderName);
                    }
                }
                newMntFolders += mnt[i] + ".zip";
            }

            // Pojedynczy mntek
            if (tempMntFile != "")
            {
                newMntFolders += tempMntFile + ".mnt";
                string curFile = tempMntFile + ".mnt";

                var folderPath = Directory.GetDirectories(PrerequisitiesPath + @"Base MNT\" + tempBrandShortcut, "City", SearchOption.AllDirectories);
                bool checkIfSingleMntExist = false;
                foreach (string file in Directory.GetFiles(folderPath[0]))
                {
                    if (Path.GetFileName(file) == curFile)
                    {
                        Console.WriteLine("Istnieje pojedynczy mnt " + tempMntFile + ".mnt");
                        checkIfSingleMntExist = true;
                        break;
                    }
                }
                if (!checkIfSingleMntExist)
                {
                    Console.WriteLine("Nie moze znalezc pojedynczego mnt. Pewnie z nazwa cos nie halo");
                    Console.ReadKey();
                    Environment.Exit(0);
                }
            }

            //Montreal footer
            if (tempMontrealFooter != "")
            {
                newMntFolders += tempMontrealFooter + ".zip";

                var folderPath = PrerequisitiesPath + @"\Base MNT\" + brand + @"\Country\CA\" + tempMontrealFooter;

                if (!Directory.Exists(folderPath))
                {
                    Console.WriteLine("Brakuje folderu montreal footer");
                    Console.ReadLine();
                    Environment.Exit(0);
                }
            }

            File.AppendAllText(AMERTxtFile, "Actual_MNT=" + newMntFolders);
        }

 
        
        
        
          

        public static void AMERtoTXT()
        {
            int i = 10;
            string excelFile = PrerequisitiesPath + AMERROLLOUT;
            string code;

            string baseMNT = "";
            country = "";
            string numberOfRegisters = "";
            city = "";
            List<string> list = new List<string>();

            System.Globalization.CultureInfo oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            Excel.Application excel = new Excel.Application();
            Excel.Workbook wb = excel.Workbooks.Open(excelFile);

            code = getValueFromExcelField(wb, 1, i, 6);

            try
            {
                while (storeCode!= null)
                {
                    if (storeCode == code)
                    {
                        numberOfRegisters = getValueFromExcelField(wb, 1, i, 9);
                        storeType = getValueFromExcelField(wb, 1, i, 8);
                        brand = getValueFromExcelField(wb, 1, i, 5);
                        baseMNT = getValueFromExcelField(wb, 1, i, 71);
                        baseMNT = RemoveSpace(baseMNT);
                        country = getValueFromExcelField(wb, 1, i, 2);
                        city = getValueFromExcelField(wb, 1, i, 4);
                        state = getValueFromExcelField(wb, 1, i, 3);

                        list.Add("storeCode=" + storeCode);
                        list.Add("storeType=" + storeType);
                        list.Add("brand=" + brand);
                        list.Add("city=" + city);
                        list.Add("country/state=" + country + "/" + state);
                        list.Add("numberOfRegisters=" + numberOfRegisters);
                        list.Add("MNT=" + baseMNT);

                        break;
                    }
                    i++;
                    code = getValueFromExcelField(wb, 1, i, 6);

                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("Prawdopodobnie brak stora w Amer Rollout.");
            }
            string[] lines = list.ToArray();
            System.IO.File.WriteAllLines(PrerequisitiesPath + "AMER.txt", lines);

            GetActualMNTFolderNames();
        }

        // Gets the basic information about the store provided. It is necessary to know the number of registers, mnt folder names from excel sheet and type of store.
        public static void GetBasicStoreData()
        {
            int i = 10;
            string rolloutExcelFile = PrerequisitiesPath + AMERROLLOUT;
            string baseMNT = "";
            //string country = "";
            string city = "";
            
            System.Globalization.CultureInfo oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            Excel.Application excel = new Excel.Application();
            Excel.Workbook wb = excel.Workbooks.Open(rolloutExcelFile);

            string tempStoreCode = getValueFromExcelField(wb, 1, i, 6);

            try
            {
                while (tempStoreCode != null)
                {
                    if (storeCode == tempStoreCode)
                    {
                        numberOfRegisters = getValueFromExcelField(wb, 1, i, 9);
                        storeType = getValueFromExcelField(wb, 1, i, 8);
                        brand = getValueFromExcelField(wb, 1, i, 5);
                        baseMNT = getValueFromExcelField(wb, 1, i, 71);
                        baseMNT = RemoveSpace(baseMNT);
                        country = getValueFromExcelField(wb, 1, i, 2);
                        city = getValueFromExcelField(wb, 1, i, 4);
                        state = getValueFromExcelField(wb, 1, i, 3);
                        break;
                    }

                    i++;
                    tempStoreCode = getValueFromExcelField(wb, 1, i, 6);

                }
                killExcelProc();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message + "   " + e.StackTrace);
                Console.WriteLine("Exception. Press any key");
                killExcelProc();
                Console.ReadKey();
                Environment.Exit(0);
            }
        }

        private static int getCountryStartIndex(Excel.Workbook workBook, int start, string text, int sheet, int column)
        {

            System.Globalization.CultureInfo oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            string tempText = getValueFromExcelField(workBook, sheet, start, column);

            while (tempText != text && tempText!=null)
            {
                start++;
                tempText = getValueFromExcelField(workBook, sheet, start, column);
            }

            return start;
        }

        private static int getCountryEndIndex(Excel.Workbook workBook, int end, string text, int sheet, int column)
        {

            System.Globalization.CultureInfo oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            string tempText = getValueFromExcelField(workBook, sheet, end, column);

            while (tempText == text)
            {
                end++;
                tempText = getValueFromExcelField(workBook, sheet, end, column);
            }

            return end;
        }

        private static List<int> getPropertyIndexes(Excel.Workbook wb, int start, int end, string text, int sheet, int column)
        {
            List<int> indexes = new List<int>();
            string tempType;
            string tempState;

            for(int i=start;i<end;i++)
            {
                tempType = getValueFromExcelField(wb, sheet, i, 2);
                tempState = getValueFromExcelField(wb, sheet, i, 3);

                if((tempType == "ALL" || tempType == storeType))
                {
                    if(state != "CA" && tempState=="CALIFORNIA")
                    {
                        continue;
                    }

                    indexes.Add(i);
                }
            }

            return indexes;
        }

        private static List<string> getXstorePropertiesList()
        {
            string excelFile = PrerequisitiesPath + XSTOREPROPS;
            string prop = "";
            int countryStartRow;
            int countryEndRow;
            string[] stringSeparator = new string[] {"\n" };
            string[] separatedProps;

            List<int> indexes = new List<int>();
            List<string> properties = new List<string>();

            Brand br = BrandParser.ParseEnum<Brand>(brand.ToLower());
            brandSheet = (int)br;

            Excel.Application excel = new Excel.Application();
            Excel.Workbook wb = excel.Workbooks.Open(excelFile);

            countryStartRow = getCountryStartIndex(wb, 2, country, brandSheet, 1);
            countryEndRow = getCountryEndIndex(wb, countryStartRow, country, brandSheet, 1);


            indexes = getPropertyIndexes(wb, countryStartRow, countryEndRow, storeType, brandSheet, 2);

            foreach(int el in indexes)
            {
                prop = getValueFromExcelField(wb, brandSheet, el, 7);
                prop = xStorePropertyEditor(prop);
                

                if (prop.Contains("\n"))
                {
                    separatedProps = prop.Split(stringSeparator, StringSplitOptions.RemoveEmptyEntries);
                    for(int i=0;i<separatedProps.Length;i++)
                    {
                        separatedProps[i] = xStorePropertyEditor(separatedProps[i]);
                    }
                    properties.AddRange(separatedProps);
                    continue;
                }
                if(!PropertyExistsCase(prop, properties))
                {
                    properties.Add(prop);
                }
               
            }

            properties = xStorePropertiesListEditor(properties);
            return properties;
        }

        private static string xStorePropertyEditor(string prop)
        {
            if (prop.Contains(" ") && !prop.Contains("ker.helpdesk.message") && !prop.Contains("ker.helpdesk.number"))
            {
                prop = prop.Replace(" ", "");
            }

            for (int j = 0; j < prop.Length; j++)
            {
                if (char.IsLetter(prop[j]))
                {
                    prop = prop.Substring(j);
                    break;
                }
            }

            return prop;
        }

        public static List<string> xStorePropertiesListEditor(List<string> list)
        {
            for(int i=0;i<list.Count;i++)
            {
                list[i] = xStorePropFiller(list[i], "dtv.auth.ajbhost.terminalid=", currentRegisterNumber.ToString());
                list[i] = xStorePropFiller(list[i], "dtv.auth.ajbhost.storenumber=", storeCode.ToString());
                list[i] = AJBIPPropFiller(list[i]);
            }
 
            return list;
        }

        private static bool PropertyExistsCase(string prop, List<string> list)
        {
            string temp = prop.Substring(0, prop.IndexOf("="));
            for (int j = 0; j < list.Count; j++)
            {
                if (list[j].Contains(temp))
                {
                    list[j] = prop;
                    return true;
                }
            }
            return false;
        }

        private static string xStorePropFiller(string prop, string key, string value)
        {
            if (prop.Contains(key))
            {
                prop = key + value;
            }

            return prop;
        }

        private static string AJBIPPropFiller(string prop)
        {
            

            if(prop.Contains("AJB"))
            {
                prop = prop.Replace("[AJBIP]", AJBIP);
            }

            return prop;
        }

        public static void XStorePropertiesToTxt(int registerNumber)
        {
            List<string> list = new List<string>();

            list = getXstorePropertiesList();

            killExcelProc();
            System.IO.File.WriteAllLines(PrerequisitiesPath + "XstoreProp.txt", list.ToArray());
        }

        public static void killExcelProc()
        {
            foreach (Process proc in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
            {
                proc.Kill();
            }
        }
    }
}