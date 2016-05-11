using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.IO;
using System.Net.Mail;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Threading;
using System.Runtime.InteropServices;
using System.Text;

namespace XStoreInstallation
{
    class SubMethods
    {
        #region variables
        //Main folder
        private static string PrerequisitiesPath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName + @"\Prerequisites\";

        // Necessary. Files in or paths to folders in the 'Prerequisities' folder
        private string xEnvironmentFolder = @"Xenvironment_4.5\";
        private string ccencFolder = @"\CCENC files\";
        private string patchFolder = @"\Patches\";
        private string baseMntFolder = @"Base MNT\";
        private string MNTProdFolder = @"MNTProd\";
        const string ENVSHORT = "EnvironmentShortcut.bat";
        const string ENVSHORT_REG = "EnvironmentShortcutANDRegisterEdit.bat";
        private const string POSVERSION = "xstore-4.5.1.5-5.22.0-0.0-GAI-pos-install(patched2).jar";
        const string INSTALLATIONCONFIG = "XStoreInstallationConfig.ini";

        //path to other files/folders
        private string envPath = "";
        private const string XSTORE_MNT_PATH = @"C:\Xstore\download\";
        private string antPath;
        private string antFileName;

        //Properties
        private string storeType;
        private string storeCode;
        private string registerNumber;
        private string numberOfRegisters;
        private string brand;
        private string city;
        private string country;
        private string mntFromExcel;
        private string ajbIP;
        private string envType = "";

        //property files
        private string userPropS1Path = PrerequisitiesPath + "USERPropS1.txt";
        private string userPropS2Path = PrerequisitiesPath + "USERPropS2.txt";
        private string userPropS3Path = PrerequisitiesPath + "USERPropS3.txt";
        private string amerPropPath = PrerequisitiesPath + "AMER.txt";
        private string xStorePropPath = PrerequisitiesPath + "XstoreProp.txt";

        //Installation config
        public string XStoreInstall { get; set; }
        public string EnvironmnentInstall { get; set; }
        public string CommaInstall { get; set; }
        public string CherryInstall { get; set; }
        public string MNTLoader { get; set; }
        public string PrinterDriversInstall { get; set; }
        public string LoadProdMNT { get; set; }
        private string loadBaseMNT = "";
        private string sqlInstall = "";
        public string LockDowns { get; set; }

        private bool IsPrinter = false;
        private bool IsComma = false;
        private bool IsCherry = false;
        private bool IsLockdown = false;
        #endregion


        //IsCorrect()
        //cutAfterEqual(string s)
        //cutBeforeEqual(string s)
        #region private_sub_methods  
        private void IsCorrect()
        {
            Console.WriteLine("Correct? y/n");
            if (Console.ReadLine() == "y")
            { }
            else
            {
                Environment.Exit(0);
            }
        }

        // Substring by "="
        private static string cutAfterEqual(string s)
        {
            int l = s.IndexOf("=");
            if (l > 0)
            {
                return s.Substring(l + 1, s.Length - l - 1);
            }
            return "";
        }

        // Substring by "=" - including "="
        private static string cutBeforeEqual(string s)
        {
            int l = s.IndexOf("=");
            if (l > 0)
            {
                return s.Substring(0, l + 1);
            }
            return "";
        }

        private void CreateSqlUser()
        {
            try
            {
                string sqlConnectionString = @"Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=master;Data Source=localhost";
                string script = File.ReadAllText(PrerequisitiesPath + @"SQLScript\sqlSrvInit.sql");

                using (SqlConnection connection = new SqlConnection(sqlConnectionString))
                {
                    SqlCommand command = new SqlCommand(script, connection);
                    connection.Open();
                    try
                    {
                        SqlDataReader reader = command.ExecuteReader();
                    }
                    catch (Exception ex)
                    {
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("Probably sql has not been installed ;(");
                Environment.Exit(0);
            }
        }
        #endregion


        //SetStoreCode() 
        //SetStoreType() 
        //SetEvironmentType()
        #region set_properties

        // Gets necessary data from txt property file
        public void SetStoreCode()
        {
            string txtFile = PrerequisitiesPath + "USERPropS1.txt";
            var loadedTxt = File.ReadAllLines(txtFile);

            storeCode = cutAfterEqual(loadedTxt[3]);
            registerNumber = cutAfterEqual(loadedTxt[1]);
            ajbIP = cutAfterEqual(loadedTxt[5]);
        }

        // Gets necessary data from posRollout file.
        public void SetStoreType()
        {
            Console.WriteLine("Getting necessary data about the shop");

            string txtFile = PrerequisitiesPath + "AMER.txt";
            var loadedTxt = File.ReadAllLines(txtFile);
            storeType = cutAfterEqual(loadedTxt[1]);
            brand = cutAfterEqual(loadedTxt[2]);
            city = cutAfterEqual(loadedTxt[3]);
            country = cutAfterEqual(loadedTxt[4]);
            numberOfRegisters = cutAfterEqual(loadedTxt[5]);
            mntFromExcel = cutAfterEqual(loadedTxt[7]);

            Console.WriteLine("\n ################### \n");
        }

        public void SetEvironmentType()
        {
            if (storeType == "DOS" || storeType == "OUTLET/DOS")
            {
                envType = "environment_delay_full_1.7.exe";
            }
            else if (storeType == "SIS")
            {
                envType = "environment_nodelay_full_1.7.exe";
            }
        }
        #endregion


        //InstallationFilesCheck()
        //TestParameters()
        //postInstallationCheck()
        #region checks

        protected bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }


        private void EnvironmentFileCheck()
        {
            string envPath = PrerequisitiesPath + xEnvironmentFolder + envType;

            if (EnvironmnentInstall.Contains("true"))
            {
                while (IsFileLocked(new FileInfo(envPath)))
                {
                    Console.WriteLine("Missing Environment installation file");
                    Console.WriteLine("Additional check in 30s");
                    Thread.Sleep(30000);
                }
            }
        }

        private void PosFileCheck()
        {
            string posPath = @"C:\dtvinst\" + POSVERSION;

            if (XStoreInstall.Contains("true"))
            {
                while (IsFileLocked(new FileInfo(posPath)))
                {
                    Console.WriteLine("Missing XStore installation file");
                    Console.WriteLine(POSVERSION);
                    Console.WriteLine("Additional check in 30s");
                    Thread.Sleep(30000);
                }
            }
        }

        private void MntProdFilesCheck()
        {
            string MNTProdPath = PrerequisitiesPath + MNTProdFolder;

            while (!Directory.Exists(MNTProdPath))
            {
                Console.WriteLine("Missing MNT from prod");
                Console.WriteLine("Additional check in 30s");
                Thread.Sleep(30000);
            }

        }

        //Check if the parameters are correct
        public void TestParameters()
        {
            Console.WriteLine("Check loaded parameters from the excel file");
            Console.WriteLine("storeType = {0}", storeType);
            Console.WriteLine("storeCode = {0}", storeCode);
            Console.WriteLine("registerNumber = {0}", registerNumber);
            Console.WriteLine("brand = {0}", brand);
            Console.WriteLine("city = {0}", city);
            Console.WriteLine("country = {0}", country);
            Console.WriteLine("Environment type = {0}", envType);
            Console.WriteLine("ajbIP = {0}", ajbIP);

            IsCorrect();
        }

        public void postInstallationCheck()
        {
            Console.WriteLine("PostInstallationCheck");

            Dictionary<string, string> systemProperties = new Dictionary<string, string>();

            string txtFile = PrerequisitiesPath + "USERPropS3.txt";
            var loadedTxtFile = File.ReadAllLines(txtFile);

            for (int i = 0; i < loadedTxtFile.Length; i++)
            {
                systemProperties[cutBeforeEqual(loadedTxtFile[i])] = cutAfterEqual(loadedTxtFile[i]);
            }

            string baseXstoreProperties = @"C:\Xstore\updates\base-Xstore.properties";

            foreach (var item in systemProperties)
            {
                var lineNumber = File.ReadAllLines(baseXstoreProperties).Select((text, index) => new { text, line = index + 1 })
                                                                    .Where(x => x.text.Contains(item.Key)).FirstOrDefault();
                var loadedTxt = File.ReadAllLines(baseXstoreProperties);
                loadedTxt[lineNumber.line - 1] = item.Key + item.Value;
                File.WriteAllLines(baseXstoreProperties, loadedTxt);
            }

            //IsCorrect();
            Console.WriteLine("\n ################### \n");
        }

        #endregion


        //EditSystemProperties(string systemTxtPath)
        //EditAntInstall()
        //EditXstoreProperties()
        #region editing_property_files

        private void EditSystemProperties(string systemTxtPath)
        {
            Console.WriteLine("Editing system.properties");

            Dictionary<string, string> systemProperties = new Dictionary<string, string>();

            // 4 wartosci do environmenta
            var loadedTxtFile = File.ReadAllLines(userPropS1Path);

            for (int i = 0; i < 4; i++)
            {
                systemProperties[cutBeforeEqual(loadedTxtFile[i])] = cutAfterEqual(loadedTxtFile[i]);
            }

            foreach (var item in systemProperties)
            {
                var lineNumber = File.ReadAllLines(systemTxtPath).Select((text, index) => new { text, line = index + 1 })
                                                                    .Where(x => x.text.Contains(item.Key)).FirstOrDefault();
                var loadedTxt = File.ReadAllLines(systemTxtPath);
                loadedTxt[lineNumber.line - 1] = item.Key + item.Value;
                File.WriteAllLines(systemTxtPath, loadedTxt);
            }
        }

        private void EditAntInstall()
        {
            // 2. Edycja anta
            Console.WriteLine("Renaming to ant.install.properties");
            //string tempAntFileTxtPath = antPath + antFileName.Replace("properties", "txt");
            File.Move(@"C:\dtvinst\" + antFileName, @"C:\dtvinst\" + "ant.install.properties");
            Console.WriteLine("Editing ant.install");
            ///////////////////////////////////////////////////////////////////////////////////////////////////
            Dictionary<string, string> systemProperties = new Dictionary<string, string>();

            // 4 wartosci do environmenta

            var loadedTxtFile = File.ReadAllLines(userPropS2Path);

            for (int j = 0; j < loadedTxtFile.Length; j++)
            {
                systemProperties[cutBeforeEqual(loadedTxtFile[j])] = cutAfterEqual(loadedTxtFile[j]);
            }
            // Mozliwe ze zbedne?
            string antPathFileName = @"C:\dtvinst\" + "ant.install.properties";

            foreach (var item in systemProperties)
            {
                var lineNumber = File.ReadAllLines(antPathFileName).Select((text, index) => new { text, line = index + 1 })
                                                                    .Where(x => x.text.Contains(item.Key)).FirstOrDefault();
                var loadedTxt = File.ReadAllLines(antPathFileName);
                loadedTxt[lineNumber.line - 1] = item.Key + item.Value;
                File.WriteAllLines(antPathFileName, loadedTxt);
            }
            ////////////////////////////////////////////////////////////////////////////////////////////////////

            //IsCorrect();
            Console.WriteLine("\n ################### \n");
        }

        private void EditXstoreProperties()
        {
            // //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // EDIT AND RUN CONFIGURE.BAT
            Console.WriteLine("Editing xstore.properties based on Xstore_properties_configuration_checks4.5+_AMER_v5_DRAFT");

            //System.IO.File.Move(@"C:\xstore\updates\xstore.properties", xStorePropTxtFile);

            Dictionary<string, string> systemProperties = new Dictionary<string, string>();
            string[] loadedTxtFile = File.ReadAllLines(xStorePropPath);

            for (int j = 0; j < loadedTxtFile.Length; j++)
            {
                systemProperties[cutBeforeEqual(loadedTxtFile[j])] = cutAfterEqual(loadedTxtFile[j]);
            }

            string xStorePropFile = @"C:\xstore\updates\xstore.properties";

            foreach (var item in systemProperties)
            {
                var lineNumber = File.ReadAllLines(xStorePropFile).Select((text, index) => new { text, line = index + 1 })
                                                                    .Where(x => x.text.Contains(item.Key)).FirstOrDefault();
                var loadedTxt = File.ReadAllLines(xStorePropFile);
                if (lineNumber == null)
                {
                    File.AppendAllText(xStorePropFile, "\n" + item.Key + item.Value);
                }
                else
                {
                    loadedTxt[lineNumber.line - 1] = item.Key + item.Value;
                    File.WriteAllLines(xStorePropFile, loadedTxt);
                }
            }

            //IsCorrect();
            Console.WriteLine("\n ################### \n");
        }
        #endregion


        //EnvironmentInstall()
        //XstoreInstall()
        //CommaInstall()
        //CherryKeyboardInstall()
        //InstallPrinterDrivers()
        #region installation

        public void Install()
        {

            Process.Start(PrerequisitiesPath + "UnZip.bat").WaitForExit();

            if (PrinterDriversInstall.Contains("true"))
            {
                PrinterDriverInstall();
            }

            if (CherryInstall.Contains("true"))
            {
                CherryKeyboardInstall();
            }

            if (EnvironmnentInstall.Contains("true"))
            {
                EnvironmentFileCheck();
                EnvironmentInstall();
            }

            // new Thread(new ThreadStart(CommaPrinterCherry)).Start();

            if (sqlInstall.Contains("true"))
            {
                SqlInstall();
            }

            if (XStoreInstall.Contains("true"))
            {
                PosFileCheck();
                XstoreInstall();
                postInstallationCheck();
            }

            if (LoadProdMNT.Contains("true"))
            {
                MntProdFilesCheck();
                CopyProdMNT();
            }

            if (loadBaseMNT.Contains("true"))
            {
                CopyBaseMNT();
            }

            if (MNTLoader.Contains("true"))
            {
                ExecuteDataLoader();
            }

            CommaPrinterCherry();

            if (LockDowns.Contains("true"))
            {
                if (storeType.ToLower() == "sis")
                {
                    Process.Start(PrerequisitiesPath + ENVSHORT);
                }
                else
                {
                    Process.Start(PrerequisitiesPath + ENVSHORT_REG);
                }
                IsLockdown = true;
            }

            QuickSummary();
        }

        private void CommaPrinterCherry()
        {


            if (CommaInstall.Contains("true"))
            {
                CommaInstallation();
            }


        }

        public static void DeepCopy(DirectoryInfo source, DirectoryInfo target)
        {

            // Recursively call the DeepCopy Method for each Directory
            foreach (DirectoryInfo dir in source.GetDirectories())
                DeepCopy(dir, target.CreateSubdirectory(dir.Name));

            // Go ahead and copy each file in "source" to the "target" directory
            foreach (FileInfo file in source.GetFiles())
                file.CopyTo(Path.Combine(target.FullName, file.Name), true);

        }

        public void SqlInstall()
        {
            string sqlBatFilePath = PrerequisitiesPath + "SqlInstaller.bat";
            string BMCSqlServerFilePath = @"C:\POS-Software-Sources\DTVINST\SQLEXPRWT_x64_ENU.exe";
            string SqlDestPath = PrerequisitiesPath + @"SQLScript\";

            string copyCommandInSqlBat = "xcopy /s " + BMCSqlServerFilePath + " \"" + SqlDestPath + "\" /y";
            var loadedTxt = File.ReadAllLines(sqlBatFilePath);
            // SUPER WAZNE CYFRY LINIJEK
            // =====================================================
            loadedTxt[1] = copyCommandInSqlBat;
            loadedTxt[2] = "cd " + PrerequisitiesPath + "SQLScript";
            // =====================================================
            File.WriteAllLines(sqlBatFilePath, loadedTxt);
            System.Diagnostics.Process.Start(sqlBatFilePath).WaitForExit();

            CreateSqlUser();
        }


        // Runs the installation of Environment depending on storeType
        public void EnvironmentInstall()
        {
            Console.WriteLine("Environment Installation");

            envPath = PrerequisitiesPath + xEnvironmentFolder + envType;

            Process.Start(envPath).WaitForExit();

            if (registerNumber == "1")
            {
                Console.WriteLine("Deleting NONLEAD, renaming LEAD");
                File.Delete(@"C:\environment\NONLEAD.system.properties");
                System.IO.File.Move(@"C:\environment\LEAD.system.properties", @"C:\environment\system.properties");
                EditSystemProperties(@"C:\environment\system.properties");
            }
            else
            {
                Console.WriteLine("Deleting LEAD, renaming NONLEAD");
                File.Delete(@"C:\environment\LEAD.system.properties");
                System.IO.File.Move(@"C:\environment\NONLEAD.system.properties", @"C:\environment\system.properties");
                EditSystemProperties(@"C:\environment\system.properties");
            }

            //IsCorrect();
            Console.WriteLine("\n ################### \n");
        }

        public void XstoreInstall()
        {
            Console.WriteLine("Xstore Installation");
            // 1. POS musi być w c:/dtvinst

            // 2. Przerzut anta z base MNT
            Console.WriteLine("Copying ant file");
            string tempStoreTypeFolderName = storeType;
            string shortcutTempStoreTypeFolderName = storeType;
            if (brand.ToLower() == "gucci")
            {
                if (storeType.ToLower() == "outlet/dos")
                {
                    tempStoreTypeFolderName = "outlet";
                    shortcutTempStoreTypeFolderName = "out";
                }
            }
            if (brand.ToLower() == "slp")
            {
                if (storeType.ToLower() == "outlet/dos")
                {
                    tempStoreTypeFolderName = "outlet";
                    shortcutTempStoreTypeFolderName = "out";
                }
            }
            if (brand.ToLower() == "bal")
            {
                if (storeType.ToLower() == "outlet/dos")
                {
                    tempStoreTypeFolderName = "outlet";
                    shortcutTempStoreTypeFolderName = "out";
                }
            }
            if (brand.ToLower() == "amq")
            {
                if (storeType.ToLower() == "outlet/dos")
                {
                    tempStoreTypeFolderName = "outlet";
                    shortcutTempStoreTypeFolderName = "out";
                }
            }
            if (brand.ToLower() == "bv")
            {
                if (storeType.ToLower() == "outlet/dos")
                {
                    tempStoreTypeFolderName = "outlet";
                    shortcutTempStoreTypeFolderName = "out";
                }
            }

            if (brand.ToLower() == "tm")
            {
                if (storeType.ToLower() == "outlet/dos")
                {
                    tempStoreTypeFolderName = "outlet";
                    shortcutTempStoreTypeFolderName = "out";
                }
            }

            if (brand.ToLower() == "sr")
            {
                if (storeType.ToLower() == "outlet/dos")
                {
                    tempStoreTypeFolderName = "outlet";
                    shortcutTempStoreTypeFolderName = "out";
                }
            }

            string prefix = "";

            switch (brand.ToLower())
            {
                case "gucci":
                    {
                        prefix = "GG";
                        break;
                    }
                case "slp":
                    {
                        prefix = "SLP";
                        break;
                    }
                case "tm":
                    {
                        prefix = "TM";
                        break;
                    }
                case "amq":
                    {
                        prefix = "AMQ";
                        break;
                    }
                case "balenciaga":
                    {
                        prefix = "BAL";
                        break;
                    }
                case "bal":
                    {
                        prefix = "BAL";
                        break;
                    }
                case "bv":
                    {
                        prefix = "BV";
                        break;
                    }
                case "sr":
                    {
                        prefix = "SR";
                        break;
                    }
                default:
                    {
                        Console.WriteLine("Cannot create prefix");
                        break;
                    }

            }
            var countryPrefix = "";
            if (country.Contains("USA"))
            {
                countryPrefix = "US";
            }
            else if (country.Contains("CANADA"))
            {
                countryPrefix = "CA";
            }
            else
            {
                countryPrefix = "MX";
            }

            // ZMIANA JAK DOJDA NOWE !!!
            if (brand.ToLower() == "slp" || brand.ToLower() == "bv" || brand.ToLower() == "gucci")
            {
                antPath = PrerequisitiesPath + @"Base MNT\" + brand + @"\Country\" + countryPrefix + @"\" + tempStoreTypeFolderName + @"\";
            }
            else
            {
                antPath = PrerequisitiesPath + @"Base MNT\" + brand + @"\Country\" + tempStoreTypeFolderName + @"\";
            }

            antFileName = "";

            switch (registerNumber)
            {
                case "1":
                    {
                        if (numberOfRegisters == "1")
                        {
                            antFileName = prefix + "_COUNTRY_" + countryPrefix + "_" + shortcutTempStoreTypeFolderName + "_SINGLEREG_ant.install.properties";
                        }
                        else
                        {
                            antFileName = prefix + "_COUNTRY_" + countryPrefix + "_" + shortcutTempStoreTypeFolderName + "_REG1_ant.install.properties";
                        }
                        break;
                    }
                case "2":
                    {
                        antFileName = prefix + "_COUNTRY_" + countryPrefix + "_" + shortcutTempStoreTypeFolderName + "_REG2_ant.install.properties";
                        break;
                    }
                default:
                    {
                        antFileName = prefix + "_COUNTRY_" + countryPrefix + "_" + shortcutTempStoreTypeFolderName + "_REGn_ant.install.properties";
                        break;
                    }
            }

            string foundAntPath = antPath + antFileName;

            if (File.Exists(@"C:\dtvinst\ant.install.properties"))
            {
                File.Delete(@"C:\dtvinst\ant.install.properties");
            }

            File.Copy(foundAntPath, @"C:\dtvinst\" + antFileName);

            Console.WriteLine("Loc: antPath: " + antPath);
            Console.WriteLine("Ant file name: " + antFileName);

            //IsCorrect();
            Console.WriteLine("\n ################### \n");


            EditAntInstall();


            // EXECUTE JARS
            Console.WriteLine("Executing pos.jar");
            System.Diagnostics.Process.Start(PrerequisitiesPath + "POS.bat").WaitForExit();
            Thread.Sleep(30000);
            // Deleting xstore/tmp files
            Console.WriteLine("Deleting anchor files (xstore/tmp)");
            foreach (string file in Directory.GetFiles(@"C:\xstore\tmp"))
            {
                if (Path.GetFileName(file).Contains("anchor"))
                {
                    File.Delete(@"C:\xstore\tmp\" + Path.GetFileName(file));
                }
            }

            

            // CCENC FILES
            Console.WriteLine("Copyings cc files");
            foreach (string file in Directory.GetFiles(PrerequisitiesPath + ccencFolder))
            {
                File.Copy(file, @"C:\xstore\res\keys\" + Path.GetFileName(file));
            }

            // Copies patches to xstore patch path... maslo maslane 
            Patching();

            //IsCorrect();
            Console.WriteLine("\n ################### \n");

            EditXstoreProperties();

            // Delete Anchor Files
            Console.WriteLine("Deleting anchor files (xstore/tmp)");
            foreach (string file in Directory.GetFiles(@"C:\xstore\tmp"))
            {
                File.Delete(@"C:\xstore\tmp\" + Path.GetFileName(file));
            }
            Console.WriteLine("2 minutes delay after delete anchor files");
            Thread.Sleep(120000);

            Console.WriteLine("Run config.bat");
            do
            {
                Process.Start(@"C:\xstore\configure.bat").WaitForExit();
            } while (checkIfGaiPosJarExists());
            Console.WriteLine("gaipos.jar does not exist");

            Console.WriteLine("Done");
            Console.WriteLine("\n ################### \n");
        }




        public void CommaInstallation()
        {
            Console.WriteLine("Installing Comma...");
            Process.Start(@"C:\dtvinst\Comma32\Comma32_SETUP\setup.exe").WaitForExit();
            Console.WriteLine("Copying Gucci directory...");

            DeepCopy(new DirectoryInfo(@"C:\dtvinst\Comma32\Gucci"), new DirectoryInfo(@"C:\Micros\Comma32\Gucci"));
            Process.Start(@"C:\Micros\Comma32\bin\Config32.exe").WaitForExit();

            Console.WriteLine("Comma has been installed.");
            //Console.ReadKey();
            IsComma = true;
        }

        public void CherryKeyboardInstall()
        {
            Console.WriteLine("Installing Cherry...");
            Process.Start(@"C:\dtvinst\cherry\DVSETUP.bat").WaitForExit();
            Process.Start(@"C:\dtvinst\cherry\program-cherry-keyboard.bat").WaitForExit();
            Process.Start(@"C:\Program Files (x86)\Cherry\Designer\Designer.exe").WaitForExit();

            Console.WriteLine("Cherry has been installed.");
            //Console.ReadKey();
            IsCherry = true;
        }

        public void PrinterDriverInstall()
        {
            Console.WriteLine("\nInstalling printer drivers...");
            Process.Start(@"C:\POS-Software-Sources\DTVINST\TMUSB610a\setup.exe").WaitForExit();

            Console.WriteLine("Drivers have been installed.");
            // Console.ReadKey();
            IsPrinter = true;
        }

        #endregion


        public void LoadConfig()
        {
            string txtFile = PrerequisitiesPath + "XStoreInstallationConfig.ini";
            var loadedTxt = File.ReadAllLines(txtFile);

            EnvironmnentInstall = cutAfterEqual(loadedTxt[0]);
            XStoreInstall = cutAfterEqual(loadedTxt[1]);
            MNTLoader = cutAfterEqual(loadedTxt[2]);
            CommaInstall = cutAfterEqual(loadedTxt[3]);
            CherryInstall = cutAfterEqual(loadedTxt[4]);
            PrinterDriversInstall = cutAfterEqual(loadedTxt[5]);
            LoadProdMNT = cutAfterEqual(loadedTxt[6]);
            loadBaseMNT = cutAfterEqual(loadedTxt[7]);
            sqlInstall = cutAfterEqual(loadedTxt[8]);
            LockDowns = cutAfterEqual(loadedTxt[9]);
        }

        #region MNT
        private void CopyBaseMNT()
        {
            Console.WriteLine("Copying MNT files");
            string temp = "";

            // MNT FILES   
            string[] stringSeparators = new string[] { ".zip" };
            var mnt = mntFromExcel.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries);

            // 07.03.2016
            string tempMntFile = "";
            string retardedMontrealFooter = "";
            foreach (var item in mnt)
            {
                if (item.Contains("CITY"))
                {
                    tempMntFile = item;
                    mnt = mnt.Where(val => val != item).ToArray();
                }

                if(item.Contains("GG_COUNTRY_CA_MONTREAL_FOOTER"))
                {
                    retardedMontrealFooter = item;
                    mnt = mnt.Where(val => val != item).ToArray();
                }
            }
            //


            // szukamy konkretnego folderu na podstawie MNTkow z tablicy string (linia 21)
            try
            {
                if (mnt.Length == 3)
                {
                    for (int i = 0; i < mnt.Length; i++)
                    {
                        var folderPath = Directory.GetDirectories(PrerequisitiesPath + baseMntFolder + brand, mnt[i], SearchOption.AllDirectories);

                        Console.WriteLine("Copying " + folderPath[0].ToString());
                        foreach (string file in Directory.GetFiles(folderPath[0]))
                        {
                            File.Copy(file, XSTORE_MNT_PATH + Path.GetFileName(file), true);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error z kopiowaniem mntkow. " + e.Message);
                Console.ReadKey();
            }

            try
            {
                if (tempMntFile != "")
                {
                    Console.WriteLine("Powinno skopiowac pojedynczy mnt");
                    var findSingleMntFile = Directory.GetFiles(PrerequisitiesPath + baseMntFolder + brand, tempMntFile, SearchOption.AllDirectories);
                    foreach (var item in findSingleMntFile)
                    {
                        if (Path.GetFileName(item) == tempMntFile)
                        {
                            Console.WriteLine("Copying single mnt " + Path.GetFileName(item));
                            File.Copy(item, XSTORE_MNT_PATH + Path.GetFileName(item), true);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error z pojedynczym mntkiem. " + e.Message);
                Console.ReadKey();
            }

            try
            {
                if (retardedMontrealFooter != "")
                {
                    Console.WriteLine("Powinno skopiowac MontrealFooter (2 mntki");
                    
                    var montrealFooterDirectory = PrerequisitiesPath + baseMntFolder + brand + @"\Country\CA\" + retardedMontrealFooter;

                    Console.WriteLine("Copying " + montrealFooterDirectory);
                    foreach (string file in Directory.GetFiles(montrealFooterDirectory))
                    {
                        File.Copy(file, XSTORE_MNT_PATH + Path.GetFileName(file), true);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error z pojedynczym mntkiem. " + e.Message);
                Console.ReadKey();
            }
        }

        private void CopyProdMNT()
        {
            var folderPath = Directory.GetDirectories(PrerequisitiesPath + MNTProdFolder).FirstOrDefault();
            Console.WriteLine("Copying prod MNTs " + folderPath);
            foreach (string file in Directory.GetFiles(Directory.GetDirectories(PrerequisitiesPath + MNTProdFolder).FirstOrDefault()))
            {
                File.Copy(file, XSTORE_MNT_PATH + Path.GetFileName(file), true);
            }
        }

        private void ExecuteDataLoader()
        {
            Console.WriteLine("Running loaddata.bat");
            Process.Start(@"C:\xstore\download\loaddata.bat").WaitForExit();
            if (!(File.Exists(@"C:\xstore\download\success.dat")))
            {
                Console.WriteLine("There was an error with MNT files (failure.dat). Press any key to continue ;(.");
                Console.ReadLine();
            }

            // Odpalenie probne environmentu
            Process.Start(@"C:\environment\environment.bat");
            Console.WriteLine("\nPress any key to delete tmp files in xstore and env");
            Console.ReadKey();
            Console.WriteLine("Deleting anchor files (xstore/tmp)");
            foreach (string file in Directory.GetFiles(@"C:\xstore\tmp"))
            {
                if (Path.GetFileName(file).Contains("anchor"))
                {
                    File.Delete(@"C:\xstore\tmp\" + Path.GetFileName(file));
                }
            }
            Console.WriteLine("Deleting anchor files (environment/tmp)");
            foreach (string file in Directory.GetFiles(@"C:\environment\tmp"))
            {
                if (Path.GetFileName(file).Contains("anchor"))
                {
                    File.Delete(@"C:\environment\tmp\" + Path.GetFileName(file));
                }
            }


        }
        #endregion

        public void Patching()
        {
            Console.WriteLine("Copying patch files");
            foreach (string file in Directory.GetFiles(PrerequisitiesPath + patchFolder))
            {
                File.Copy(file, @"C:\xstore\lib\patch\" + Path.GetFileName(file));
            }
        }

        public void QuickSummary()
        {
            Console.WriteLine("\nQuick summary:");
            if (registerNumber == "1")
            {
                Console.WriteLine("\n\tComma: " + IsComma);
            }
            Console.WriteLine("\n\tPrinter: " + IsPrinter);
            Console.WriteLine("\n\tCherry: " + IsCherry);
            Console.WriteLine("\n\tLockdown: " + IsLockdown);
        }

        private bool checkIfGaiPosJarExists()
        {
            if(File.Exists(@"C:\xstore\lib\gai-pos.jar.tmp"))
            {
                return true;
            }
            return false;
        }

    }
}