using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using OSIsoft.AF;
using OSIsoft.AF.Analysis;
using OSIsoft.AF.Asset;
using OSIsoft.AF.UnitsOfMeasure;
using System;
using System.Collections.Generic;
using System.ComponentModel.Composition.Primitives;
using System.Data;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Security.Policy;
using ValueKit_ExtractionUI;

namespace ConsoleApp_Tests_Xavier
{
    class Program
    {
        static void Main(string[] args)
        {
            Utilities util = new Utilities();
            util.WriteInBlue("Which Unit Kit would you like to export:");
            Console.ForegroundColor= ConsoleColor.White;           
            string unitkit = Console.ReadLine();

            string pisystem = "pre-sales-pi";
            string myDB = "Refinery Process Unit Kits";
            string username = "pre-sales-pi\\xavier.fourrage";
            string password = "Welcome123!";
            bool credentials = true;
            if (util.Confirm("Default PI System is: pre-sales-pi. AF Database is: Refinery Process Unit Kits . Would you like to rename those?"))
            {
                util.WriteInBlue("Enter the PI System name:");
                Console.ForegroundColor = ConsoleColor.White;
                pisystem = Console.ReadLine();
                util.WriteInBlue("Enter the AF Database name:");
                Console.ForegroundColor = ConsoleColor.White;
                myDB = Console.ReadLine();
                if (!util.Confirm("Connect with your current windows user?"))
                {
                    util.WriteInBlue("Username:");
                    Console.ForegroundColor = ConsoleColor.White;
                    username = Console.ReadLine();

                    util.WriteInBlue("Username password:");
                    Console.ForegroundColor = ConsoleColor.White;
                    password = Console.ReadLine();
                }
                else
                {
                    credentials=false;
                }
            };

            PISystem pisys=PIsystemConnect(pisystem, username, password,credentials);
            if (pisys != null)
            {
                util.WriteInBlue("Connected to " + pisystem);
                AFDatabase afdB = GetAFdatabase(pisystem, myDB);
                if (afdB != null)
                {
                    util.WriteInBlue("Connected to " + afdB);
                    XLWorkbook wb = new XLWorkbook();                
                    AFElement UnitKit_Search = (AFElement.FindElements(afdB, null, unitkit, AFSearchField.Name, true, AFSortField.Name, AFSortOrder.Ascending, 10000).Count > 1) ? AFElement.FindElements(afdB, null, unitkit, AFSearchField.Name, true, AFSortField.Name, AFSortOrder.Ascending, 10000)[0] : null;

                    if (UnitKit_Search != null)
                    {
                        /*******************CREATING THE UOM AF Database WORKSHEET****************/
                        List<UOM> uOM_list = GetCustomUOM(pisystem);
                        DataTable dt_UOM = UomDataTable(uOM_list);
                        wb = AddNewWorkSheetToWb(wb, dt_UOM, "UOM AF Database");
                        /**********************************************************************/

                        /*******************CREATING THE CATEGORY WORKSHEET****************/
                        AFCategories afcat = ListAttributeCategories(afdB);
                        AFCategories elcat = ListElementCategories(afdB, unitkit);
                        DataTable dt_categories = AFCategories(afcat, elcat);
                        wb = AddNewWorkSheetToWb(wb, dt_categories, "AF Categories");
                        /**********************************************************************/

                        /*******************CREATING THE ENUMERATION SETS WORKSHEET****************/
                        AFNamedCollection<AFEnumerationSet> listenumSets = ListEnumSets(afdB);
                        DataTable dt_enumSet = EnumerationSetsDataTable(listenumSets);
                        wb = AddNewWorkSheetToWb(wb, dt_enumSet, "AF Enumeration Sets");
                        /**********************************************************************/

                        /*******************CREATING THE CONFIGURATION ELEMENT WORKSHEET****************/
                        AFElement configElement = GetConfigurationElement(afdB);

                        util.WriteInBlue("Exporting the config element...");
                        ExportToXML(configElement, pisystem, "_Configuration_AFElement.xml");
                        util.WriteInBlue("_Configuration element has been exported under _Configuration_AFElement.xml");

                        DataTable dt_configElement = ConfigElementDataTable(configElement);
                        wb = AddNewWorkSheetToWb(wb, dt_configElement, "Configuration Element");
                        /**********************************************************************/

                        /*******************CREATING THE ELMENT TEMPLATES WORKSHEET****************/
                        util.WriteInBlue("Exporting the " + unitkit + " element...");
                        ExportToXML(UnitKit_Search, pisystem, unitkit + "_AFElement.xml");
                        util.WriteInBlue(unitkit + " has been exported under " + unitkit + "_AFElement.xml");

                        AFElements afelements = UnitKit_Search.Elements;
                        List<AFElement> afelList = GetElementsRecursively(afelements);
                        afelList.Insert(0, UnitKit_Search);
                        afelList.Insert(0, UnitKit_Search.Parent);
                        List<AFElementTemplate> afelTempList = new List<AFElementTemplate>();
                        foreach (AFElement afel in afelList)
                        {
                            if (GetElementTemplate(afel) != null)
                            {
                                if (!afelTempList.Contains(GetElementTemplate(afel)))
                                {
                                    afelTempList.Add(GetElementTemplate(afel));
                                    if (GetElementTemplate(afel).BaseTemplate != null && !afelTempList.Contains(GetElementTemplate(afel).BaseTemplate))
                                    {
                                        afelTempList.Add(GetElementTemplate(afel).BaseTemplate);
                                        if (GetElementTemplate(afel).BaseTemplate.BaseTemplate != null && !afelTempList.Contains(GetElementTemplate(afel).BaseTemplate.BaseTemplate))
                                        {
                                            afelTempList.Add(GetElementTemplate(afel).BaseTemplate.BaseTemplate);
                                            if (GetElementTemplate(afel).BaseTemplate.BaseTemplate.BaseTemplate != null && !afelTempList.Contains(GetElementTemplate(afel).BaseTemplate.BaseTemplate.BaseTemplate))
                                            {
                                                afelTempList.Add(GetElementTemplate(afel).BaseTemplate.BaseTemplate.BaseTemplate);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        //remove duplicates if any
                        afelTempList = afelTempList.Distinct().ToList();
                        DataTable dt_elTemp = ElementTemplateListDataTable(Order_AFElementTemplateList(afelTempList));
                        wb = AddNewWorkSheetToWb(wb, dt_elTemp, "Element Templates");
                        /**********************************************************************/

                        /*******************CREATING THE ELEMENTS WORKSHEET****************/
                        DataTable dt_elements = ElementsDataTable(afelList, pisystem, myDB);
                        wb = AddNewWorkSheetToWb(wb, dt_elements, "Element");
                        /**********************************************************************/

                        /*******************CREATING THE EF TEMPLATES WORKSHEET****************/
                        List<AFElementTemplate> aFEventFrameTemplates = GetEFTemplateList(afdB);
                        DataTable dt_EFTemplate = EventFrameTemplateDataTable(Order_AFElementTemplateList(aFEventFrameTemplates));
                        wb = AddNewWorkSheetToWb(wb, dt_EFTemplate, "EF Template");
                        /**********************************************************************/

                        /*******************CREATING THE AF TABLE WORKSHEETS****************/
                        List<AFTable> aFTables = GetAFTablesList(afelList, afdB);
                        foreach (AFTable aFTable in aFTables)
                        {
                            DataTable dt_table = DataTable_AFTable(aFTable);
                            string sheetname = aFTable.Name;
                            wb = AddNewWorkSheetToWb(wb, dt_table, sheetname);
                        }
                        /**********************************************************************/

                        string date = DateTime.Now.ToString("yyyy MMMM dd_HH-mm-ss");
                        wb.SaveAs(unitkit + "_AFexport_" + date + ".xlsx");
                        util.WriteInBlue("Output has been saved under: " + unitkit + "_AFexport_" + date + ".xlsx");
                        util.PressEnterToExit();
                    }
                    else
                    {
                        util.WriteInRed(unitkit + " element not found in " + afdB + " AF Database");
                        util.PressEnterToExit();
                    }
                }
                else
                {
                    util.WriteInRed("Could not connect to AF Database " + afdB);
                    util.PressEnterToExit();
                }
            }
            else
            {
                util.WriteInRed("Could not connect to PI System "+pisystem);
                util.PressEnterToExit();
            }
            
            
        }
        public static void ExportToXML(object configElement, string pisystem, string filename)
        {
            PISystems myPISystems = new PISystems();
            PISystem myPISystem = myPISystems[pisystem];
            System.EventHandler<AFProgressEventArgs> eventHandler = null;
            PIExportMode pIExport = PIExportMode.AllReferences | PIExportMode.NoUniqueID;
            myPISystem.ExportXml(configElement, pIExport, filename, null, null, eventHandler); ;
        }
        public static DataTable DataTable_AFTable(AFTable aFTable)
        {
            return aFTable.Table;
        }
        public static List<AFTable> GetAFTablesList(List<AFElement> afelList, AFDatabase afdb)
        {
            List<AFTable> aftableList= new List<AFTable>();

            foreach(AFElement afel in afelList)
            {               
                foreach (AFAttribute att in afel.Attributes)
                {
                    if(att.DataReferencePlugIn!=null)
                    {
                        if (att.DataReferencePlugIn.ToString() == "Table Lookup")
                        {
                            var attrConfigString = att.ConfigString;
                            int pFrom = attrConfigString.LastIndexOf("FROM")+5;
                            int pTo = attrConfigString.IndexOf(" WHERE");

                            String result = attrConfigString.Substring(pFrom, pTo - pFrom);
                            
                            if (result.First() =='[')
                            {
                                result= result.Substring(1);
                            }
                      
                            if (result.Last() == ']')
                            {
                                result = result.Remove(result.Length - 1, 1);
                            }                          
                            AFTable aftable = (AFTable.FindTables(afdb, result, AFSearchField.Name, AFSortField.Name, AFSortOrder.Descending, 100).Count>0)? 
                                              AFTable.FindTables(afdb, result, AFSearchField.Name, AFSortField.Name, AFSortOrder.Descending, 100)[0]:null;
                            if (aftable != null)
                            { aftableList.Add(aftable); }
                        }                        
                    }
                }
            }  
            aftableList=aftableList.Distinct().ToList();
            return aftableList;
        }
        public static void GetAllPropertiesFromObject(object obj)
        {
            foreach (var prop in obj.GetType().GetProperties())
            {
                Console.WriteLine("{0}={1}", prop.Name, prop.GetValue(obj, null));
            }
        }
        public static int GetDepth(List<AFElementTemplate> list_elementTemplate,AFElementTemplate ElementTemplate)
        {
            if (ElementTemplate.BaseTemplate == null) return 0;
            return GetDepth(list_elementTemplate, list_elementTemplate.Find(el=> ElementTemplate.BaseTemplate.Name==el.Name))+1;
        }
        public static List<AFElementTemplate> Order_AFElementTemplateList(List<AFElementTemplate> afelTempList)
        {
            List<Tuple<AFElementTemplate, int>> afelTempList_temp_sorted = new List<Tuple<AFElementTemplate, int>>();
            foreach (var el in afelTempList)
            {
                /*Console.WriteLine(el.Name + " depth: "+GetDepth(afelTempList, el));*/
                Tuple<AFElementTemplate, int> tpl = new Tuple<AFElementTemplate, int>(el, GetDepth(afelTempList, el));
                afelTempList_temp_sorted.Add(tpl);
            }
            afelTempList_temp_sorted = afelTempList_temp_sorted.OrderBy(el => el.Item2).ToList();
            List<AFElementTemplate> afelTempList_sorted = new List<AFElementTemplate>();
            foreach (var tpl in afelTempList_temp_sorted)
            {
                afelTempList_sorted.Add(tpl.Item1);
                /*  Console.WriteLine(tpl.Item1.Name +" ,depth: "+tpl.Item2);*/
            }
            return afelTempList_sorted;
        }
            public static XLWorkbook AddNewWorkSheetToWb(XLWorkbook wb,DataTable dt, string worksheetname)
        {
            List<int?> maximumLengthForColumns =
            Enumerable.Range(0, dt.Columns.Count)
             .Select(col => dt.AsEnumerable()
            .Select(row => row[col]).OfType<string>()
            .Max(val => val?.Length)
             ).ToList();

            var ws = wb.AddWorksheet(worksheetname);
            ws.Cell("A1").InsertData(dt);

            int colnum = dt.Columns.Count;
            /*char c1 = 'A';*/
            for (int i=0;i< dt.Columns.Count; i++)
            {
                var col = ws.Column(GetColNameFromIndex(i+1).ToString());
                col.Width = maximumLengthForColumns[i]!=null? (int)maximumLengthForColumns[i]:9;
                /*c1++;*/
            }
            return wb;
        }
        // (1 = A, 2 = B...27 = AA...703 = AAA...)
        public static string GetColNameFromIndex(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
        public static PISystem PIsystemConnect(string pisystem, string username, string password,bool credentials)
        {
            Utilities utilities= new Utilities();
            try
            {
                PISystems myPISystems = new PISystems();
                PISystem myPISystem = myPISystems[pisystem];
                // Connect using a specified credential.
                if (credentials)
                {
                    NetworkCredential credential = new NetworkCredential(username, password);
                    myPISystem.Connect(credential);
                }              
                return myPISystem;
            }
            catch (Exception ex)
            {
                // Expected exception since credential needs a valid user name and password.
                Console.WriteLine(ex.Message);
                utilities.WriteInRed("could not connect to " + pisystem);
                return null;
            }
        }
        public static List<UOM> GetCustomUOM(string pisystem)
        {
            PISystems myPISystems = new PISystems();
            PISystem myPISystem = myPISystems[pisystem]; //TO BE SWITCH TO PISYSTEM CLASS, ADD THIS AT THE TOP OF MAIN

            List<UOM> uom_list = new List<UOM>();
            foreach (UOM uom in myPISystem.UOMDatabase.UOMs)
            {
                if (uom.Origin.ToString() != "SystemDefined")
                {
                    uom_list.Add(uom);
                }
            }
            return uom_list;
        }
        public static DataTable UomDataTable(List<UOM> uom_list)
        {
            DataTable dt_uom = new DataTable();
            dt_uom.Columns.Add("Selected(x)", typeof(string));
            dt_uom.Columns.Add("Parent", typeof(string));
            dt_uom.Columns.Add("Name", typeof(string));
            dt_uom.Columns.Add("ObjectType", typeof(string));
            dt_uom.Columns.Add("Description", typeof(string));
            dt_uom.Columns.Add("Abbreviation", typeof(string));
            dt_uom.Columns.Add("Origin", typeof(string));
            dt_uom.Columns.Add("RefUOM", typeof(string));
            dt_uom.Columns.Add("RefFactor", typeof(string));
            dt_uom.Columns.Add("RefOffset", typeof(string));
            dt_uom.Columns.Add("RefFormulaTo", typeof(string));
            dt_uom.Columns.Add("[Metric]", typeof(string));
            dt_uom.Columns.Add("[US Customary]", typeof(string));

            DataRow row_init = dt_uom.NewRow();
            row_init["Selected(x)"] = "Selected(x)";
            row_init["Parent"] = "Parent";
            row_init["Name"] = "Name";
            row_init["ObjectType"] = "ObjectType";
            row_init["Description"] = "Description";
            row_init["Abbreviation"] = "Abbreviation";
            row_init["Origin"] = "Origin";
            row_init["RefUOM"] = "RefUOM";
            row_init["RefFactor"] = "RefFactor";
            row_init["RefOffset"] = "RefOffset";
            row_init["RefFormulaTo"] = "RefFormulaTo";
            row_init["[Metric]"] = "[Metric]";
            row_init["[US Customary]"] = "[US Customary]";
            dt_uom.Rows.Add(row_init);

            foreach(UOM uom in uom_list)
            {
                DataRow row = dt_uom.NewRow();
                row["Selected(x)"] = "x";
                row["Parent"] = uom.Class;
                row["Name"] = uom.Name;
                row["ObjectType"] = uom.Identity;
                row["Description"] = uom.Description;
                row["Abbreviation"] = uom.Abbreviation;
                row["Origin"] = uom.Origin;
                row["RefUOM"] = uom.RefUOM;
                row["RefFactor"] = uom.RefFactor;
                row["RefOffset"] = uom.RefOffset;
                row["RefFormulaTo"] = uom.RefFormulaTo;
                row["[Metric]"] = ""; //TO FIX
                row["[US Customary]"] = ""; // TO FIX
                dt_uom.Rows.Add(row);
            }

            return dt_uom;
        }
        public static DataTable ElementsDataTable(List<AFElement> afelList, string pisystem, string afdb)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Selected(x)", typeof(string));
            dt.Columns.Add("Parent", typeof(string));
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("ObjectType", typeof(string));
            dt.Columns.Add("Categories", typeof(string));
            dt.Columns.Add("Template", typeof(string));
            dt.Columns.Add("|Latitude", typeof(string));
            dt.Columns.Add("|Longitude", typeof(string));

            DataRow row_init = dt.NewRow();
            row_init["Selected(x)"] = "Selected(x)";
            row_init["Parent"] = "Parent";
            row_init["Name"] = "Name";
            row_init["ObjectType"] = "ObjectType";
            row_init["Categories"] = "Categories";
            row_init["Template"] = "Template";
            row_init["|Latitude"] = "|Latitude";
            row_init["|Longitude"] = "|Longitude";
            dt.Rows.Add(row_init);

            foreach(AFElement afel in afelList)
            {
                DataRow row = dt.NewRow();
                row["Selected(x)"] = "x";
                row["Parent"] = GetParentPath(afel, pisystem, afdb);
                row["Name"] = afel.Name;
                row["ObjectType"] = afel.Identity;
                row["Categories"] = afel.CategoriesString;
                row["Template"] = afel.Template;
                row["|Latitude"] = ""; //TO BE FIXED
                row["|Longitude"] = ""; // TO BE FIXED 
                dt.Rows.Add(row);
            }

            return dt;
        }
        public static string GetParentPath(AFElement el,string pisystem,string mydb)
        {
            List<AFElement> list = new List<AFElement>();
            list.Add(el);
            string relativepath = "\\\\"+pisystem+"\\"+mydb;
            IDictionary<Guid, String> keyValuePairs_list = AFElement.GetPath(relativepath, list);
            var val="";          
            foreach (KeyValuePair<Guid, string> kvp in keyValuePairs_list)
            {
              /*  Console.WriteLine("Key = {0}, Value = {1}",
                    kvp.Key, kvp.Value);*/
                val= kvp.Value;
            }
            string[] test = val.Split('\\');
            test = test.Take(test.Count() - 1).ToArray();
            var val2= string.Join("\\", test);
            return val2;
        }
        public static AFDatabase GetAFdatabase(string pisystem, string mydB)
        {
            Utilities utilities= new Utilities();
            try
            {
                PISystems myPISystems = new PISystems();
                PISystem myPISystem = myPISystems[pisystem];
                AFDatabase afdb = myPISystem.Databases[mydB];
                return afdb;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                utilities.WriteInRed("could not connect to " + mydB);
                return null;
            }
        }
        public static AFCategories ListAttributeCategories(AFDatabase AFdB)
        {
            AFCategories AFAttributeCategoriesList = AFdB.AttributeCategories;
            /*foreach (AFCategory category in AFAttributeCategoriesList) { Console.WriteLine(category.Name); }*/
            return AFAttributeCategoriesList;
        }
        public static AFCategories ListElementCategories(AFDatabase AFdB, string unitkit)
        {
            AFCategories AFElementCategoriesList = AFdB.ElementCategories;          
            for (int i= AFElementCategoriesList.Count-1; i>=0; i--)
            {
                if (!(AFElementCategoriesList[i].Name.Contains(unitkit) || AFElementCategoriesList[i].Name.Contains("Base") || AFElementCategoriesList[i].Name.Contains("Menu")))
                {
                    AFElementCategoriesList.Remove(AFElementCategoriesList[i]);
                }
            }
           /* foreach(AFCategory afcat in AFElementCategoriesList)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine(afcat.Name);
            }*/
            return AFElementCategoriesList;
        }
       
        public static DataTable AFCategories(AFCategories AttCategories, AFCategories ElCategories)
        {
            DataTable AFCategoriesList = new DataTable();
            AFCategoriesList.Columns.Add("Selected(x)", typeof(string));
            AFCategoriesList.Columns.Add("Parent", typeof(string));
            AFCategoriesList.Columns.Add("Name", typeof(string));
            AFCategoriesList.Columns.Add("ObjectType", typeof(string));

            DataRow row_init = AFCategoriesList.NewRow();
            row_init["Selected(x)"] = "Selected(x)";
            row_init["Parent"] = "Parent";
            row_init["Name"] = "Name";
            row_init["ObjectType"] = "ObjectType";
            AFCategoriesList.Rows.Add(row_init);

            foreach (AFCategory category in AttCategories)
            {
                DataRow row = AFCategoriesList.NewRow();
                row["Selected(x)"] = "x";
                row["Parent"] = "";
                row["Name"] = category.Name;
                row["ObjectType"] = category.Identity;
                AFCategoriesList.Rows.Add(row);
            }
            foreach (AFCategory category in ElCategories)
            {
                DataRow row = AFCategoriesList.NewRow();
                row["Selected(x)"] = "x";
                row["Parent"] = "";
                row["Name"] = category.Name;
                row["ObjectType"] = category.Identity;
                AFCategoriesList.Rows.Add(row);
            }
            return AFCategoriesList;
        }
      
        public static AFNamedCollection<AFEnumerationSet> ListEnumSets(AFDatabase AFdB)
        {
            AFNamedCollection<AFEnumerationSet> AFEnumSetList = AFdB.EnumerationSets;
/*            foreach (AFEnumerationSet enumset in AFEnumSetList)
            {              
                    Console.ForegroundColor = ConsoleColor.Blue;
                    Console.WriteLine(enumset.Name);
                    *//*ListEnumerationSetValues(enumset);*//*              
            }*/
        return AFEnumSetList;
        }
        public static DataTable EnumerationSetsDataTable(AFNamedCollection<AFEnumerationSet> AFEnumSetList)
        {
            DataTable AFEnumerationSetsList = new DataTable();
            AFEnumerationSetsList.Columns.Add("Selected(x)", typeof(string));
            AFEnumerationSetsList.Columns.Add("Parent", typeof(string));
            AFEnumerationSetsList.Columns.Add("Name", typeof(string));
            AFEnumerationSetsList.Columns.Add("ObjectType", typeof(string));
            AFEnumerationSetsList.Columns.Add("EnumerationValue", typeof(string));

            DataRow row_init = AFEnumerationSetsList.NewRow();
            row_init["Selected(x)"] = "Selected(x)";
            row_init["Parent"] = "Parent";
            row_init["Name"] = "Name";
            row_init["ObjectType"] = "ObjectType";
            row_init["EnumerationValue"] = "EnumerationValue";
            AFEnumerationSetsList.Rows.Add(row_init);

            foreach (AFEnumerationSet afenum in AFEnumSetList)
            {
                DataRow row = AFEnumerationSetsList.NewRow();
                row["Selected(x)"] = "x";
                row["Parent"] = "";
                row["Name"] = afenum.Name;
                row["ObjectType"] = afenum.Identity;
                row["EnumerationValue"] = "";
                AFEnumerationSetsList.Rows.Add(row);

                foreach(AFEnumerationValue afenumValue in afenum)
                {
                    DataRow row_val = AFEnumerationSetsList.NewRow();
                    row_val["Selected(x)"] = "x";
                    row_val["Parent"] = afenum.Name;
                    row_val["Name"] = afenum.GetByValue(afenumValue.Value);
                    row_val["ObjectType"] = "EnumerationValue";
                    row_val["EnumerationValue"] = afenumValue.Value;
                    AFEnumerationSetsList.Rows.Add(row_val);
                }
            }           
            return AFEnumerationSetsList;
        }
        public static void ListEnumerationSetValues(AFEnumerationSet myEnumerationSet)
        {
            foreach (AFEnumerationValue eVal in myEnumerationSet)
            {
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine("Name in set = {0}", eVal.Name);
                Console.WriteLine("Value in set = {0}", eVal.Value);
            }
        }
        public static AFElement GetConfigurationElement(AFDatabase afdb)
        {
            AFNamedCollectionList<AFElement> AFElementSearch = AFElement.FindElements(afdb, null, "_Configuration", OSIsoft.AF.AFSearchField.Name, true, OSIsoft.AF.AFSortField.Name, OSIsoft.AF.AFSortOrder.Ascending, 10000);
            return AFElementSearch[0];
        }
        public static DataTable ConfigElementDataTable(AFElement configEl)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Selected(x)", typeof(string));
            dt.Columns.Add("Parent", typeof(string));
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("ObjectType", typeof(string));
            dt.Columns.Add("AttributeIsHidden", typeof(string));
            dt.Columns.Add("AttributeIsManualDataEntry", typeof(string));
            dt.Columns.Add("AttributeTrait", typeof(string));
            dt.Columns.Add("AttributeIsConfigurationItem", typeof(string));
            dt.Columns.Add("Description", typeof(string));
            dt.Columns.Add("Categories", typeof(string));
            dt.Columns.Add("AttributeDefaultUOM", typeof(string));
            dt.Columns.Add("AttributeType", typeof(string));
            dt.Columns.Add("AttributeValue", typeof(string));
            dt.Columns.Add("AttributeDisplayDigits", typeof(string));


            DataRow row_init = dt.NewRow();
            row_init["Selected(x)"] = "Selected(x)";
            row_init["Parent"] = "Parent";
            row_init["Name"] = "Name";
            row_init["ObjectType"] = "ObjectType";
            row_init["AttributeIsHidden"] = "AttributeIsHidden";
            row_init["AttributeIsManualDataEntry"] = "AttributeIsManualDataEntry";
            row_init["AttributeTrait"] = "AttributeTrait";
            row_init["AttributeIsConfigurationItem"] = "AttributeIsConfigurationItem";
            row_init["Description"] = "Description";
            row_init["Categories"] = "Categories";
            row_init["AttributeDefaultUOM"] = "AttributeDefaultUOM";
            row_init["AttributeType"] = "AttributeType";
            row_init["AttributeValue"] = "AttributeValue";
            row_init["AttributeDisplayDigits"] = "AttributeDisplayDigits";
            dt.Rows.Add(row_init);

            DataRow row_elt = dt.NewRow();
            row_elt["Selected(x)"] = "x";
            row_elt["Parent"] = configEl.Parent;
            row_elt["Name"] =configEl.Name;
            row_elt["ObjectType"] = configEl.Identity;
            row_elt["Description"] = configEl.Description;
            row_elt["Categories"] = configEl.Categories;        
            dt.Rows.Add(row_elt);
            foreach (AFAttribute att in configEl.Attributes)
            {
                DataRow row = dt.NewRow();
                row["Selected(x)"] = "x";
                row["Parent"] = configEl.Name;
                row["Name"] = att.Name;
                row["ObjectType"] = att.Identity;
                row["AttributeIsHidden"] = att.IsHidden;
                row["AttributeIsManualDataEntry"] = att.IsManualDataEntry;
                row["AttributeTrait"] = att.Trait;
                row["AttributeIsConfigurationItem"] = att.IsConfigurationItem;
                row["Description"] = att.Description;
                row["Categories"] = att.Categories;
                row["AttributeDefaultUOM"] =att.DefaultUOM;
                row["AttributeType"] = att.Type;
                row["AttributeValue"] = att.GetValue();
                row["AttributeDisplayDigits"] = att.DisplayDigits;
                dt.Rows.Add(row);
            }
            return dt;
        }
        public static List<AFElement> GetElementsRecursively(AFElements elements)
        {
            var list = new List<AFElement>();
            foreach (var element in elements)
            {
                list.Add(element);
                if (element.HasChildren)
                {
                    list.AddRange(GetElementsRecursively(element.Elements));
                }
            }
            return list;
        }
        public static AFElementTemplate GetElementTemplate(AFElement afel)
        {
            if (afel.Template != null)
            {              
                return afel.Template;
            }
            else {
                return null;
            }

        }
        public static DataTable ElementTemplateListDataTable(List<AFElementTemplate> eltemplList)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Selected(x)", typeof(string));
            dt.Columns.Add("Parent", typeof(string));
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("ObjectType", typeof(string));
            dt.Columns.Add("Categories", typeof(string));
            dt.Columns.Add("Description", typeof(string));
            dt.Columns.Add("BaseTemplate", typeof(string));
            dt.Columns.Add("Type", typeof(string));
            dt.Columns.Add("AllowElementToExtend", typeof(string));
            dt.Columns.Add("BaseTemplateOnly", typeof(string));
            dt.Columns.Add("NamingPattern", typeof(string));
            dt.Columns.Add("AttributeConfigString", typeof(string));
            dt.Columns.Add("AttributeDefaultUOM", typeof(string));
            dt.Columns.Add("AttributeIsHidden", typeof(string));
            dt.Columns.Add("AttributeIsManualDataEntry", typeof(string));
            dt.Columns.Add("AttributeTrait", typeof(string));
            dt.Columns.Add("AttributeIsConfigurationItem", typeof(string));
            dt.Columns.Add("AttributeIsExcluded", typeof(string));
            dt.Columns.Add("AttributeIsIndexed", typeof(string));
            dt.Columns.Add("AttributeTypeQualifier", typeof(string));
            dt.Columns.Add("AttributeDefaultValue", typeof(string));
            dt.Columns.Add("AttributeDataReference", typeof(string));
            dt.Columns.Add("AttributeDisplayDigits", typeof(string));
            dt.Columns.Add("AttributeType", typeof(string));
            dt.Columns.Add("AnalysisRule", typeof(string));
            dt.Columns.Add("AnalysisRuleConfigString", typeof(string));
            dt.Columns.Add("AnalysisRuleVariableMapping", typeof(string));
            dt.Columns.Add("TimeRule", typeof(string));
            dt.Columns.Add("TimeRuleconfigString", typeof(string));

            DataRow row_init = dt.NewRow();
            row_init["Selected(x)"] = "Selected(x)";
            row_init["Parent"] = "Parent";
            row_init["Name"] = "Name";
            row_init["ObjectType"] = "ObjectType";
            row_init["Categories"] = "Categories";
            row_init["Description"] = "Description";
            row_init["BaseTemplate"] = "BaseTemplate";
            row_init["Type"] = "Type";
            row_init["AllowElementToExtend"] = "AllowElementToExtend";
            row_init["BaseTemplateOnly"] = "BaseTemplateOnly";
            row_init["NamingPattern"] = "NamingPattern";
            row_init["AttributeConfigString"] = "AttributeConfigString";
            row_init["AttributeDefaultUOM"] = "AttributeDefaultUOM";
            row_init["AttributeIsHidden"] = "AttributeIsHidden";
            row_init["AttributeIsManualDataEntry"] = "AttributeIsManualDataEntry";
            row_init["AttributeTrait"] = "AttributeTrait";
            row_init["AttributeIsConfigurationItem"] = "AttributeIsConfigurationItem";
            row_init["AttributeIsExcluded"] = "AttributeIsExcluded";
            row_init["AttributeIsIndexed"] = "AttributeIsIndexed";
            row_init["AttributeTypeQualifier"] = "AttributeTypeQualifier";
            row_init["AttributeDefaultValue"] = "AttributeDefaultValue";
            row_init["AttributeDataReference"] = "AttributeDataReference";
            row_init["AttributeDisplayDigits"] = "AttributeDisplayDigits";
            row_init["AttributeType"] = "AttributeType";
            row_init["AnalysisRule"] = "AnalysisRule";
            row_init["AnalysisRuleConfigString"] = "AnalysisRuleConfigString";
            row_init["AnalysisRuleVariableMapping"] = "AnalysisRuleVariableMapping";
            row_init["TimeRule"] = "TimeRule";
            row_init["TimeRuleconfigString"] = "TimeRuleconfigString";
            dt.Rows.Add(row_init);

            foreach(AFElementTemplate afeltemplate in eltemplList)
            {               
                DataRow row = dt.NewRow();
                row["Selected(x)"] = "x";
                row["Parent"] = "";
                row["Name"] = afeltemplate.Name;
                row["ObjectType"] = afeltemplate.Identity;
                row["Categories"] = afeltemplate.GetAllCategoriesString();
                row["Description"] = afeltemplate.Description;
                row["BaseTemplate"] = afeltemplate.BaseTemplate!=null?afeltemplate.BaseTemplate.Name:"";
                row["Type"] = afeltemplate.Type;
                row["AllowElementToExtend"] = afeltemplate.AllowElementToExtend;
                row["BaseTemplateOnly"] = afeltemplate.BaseTemplateOnly;
                row["NamingPattern"] = afeltemplate.NamingPattern;
/*              row["AttributeConfigString"] = "";
                row["AttributeDefaultUOM"] = "";
                row["AttributeIsHidden"] = "";
                row["AttributeIsManualDataEntry"] = "";
                row["AttributeTrait"] = "";
                row["AttributeIsConfigurationItem"] = "";
                row["AttributeIsExcluded"] = "";
                row["AttributeIsIndexed"] = "";
                row["AttributeTypeQualifier"] = "";
                row["AttributeDefaultValue"] = "";
                row["AttributeDataReference"] = "";
                row["AttributeDisplayDigits"] = "";
                row["AttributeType"] = "";*/
                dt.Rows.Add(row);

                foreach (AFAttributeTemplate att in afeltemplate.AttributeTemplates)
                {
                    DataRow row_att = dt.NewRow();
                    row_att["Selected(x)"] = "x";
                    row_att["Parent"] = afeltemplate.Name;
                    row_att["Name"] = att.Name;
                    row_att["ObjectType"] = att.Identity;
                    row_att["Categories"] = att.CategoriesString;
                    row_att["Description"] = att.Description;
/*                    row_att["BaseTemplate"] = "";
                    row_att["Type"] = "";
                    row_att["AllowElementToExtend"] = "";
                    row_att["BaseTemplateOnly"] = "";
                    row_att["NamingPattern"] = "";*/
                    row_att["AttributeConfigString"] = att.ConfigString;
                    row_att["AttributeDefaultUOM"] = att.DisplayUOM;
                    row_att["AttributeIsHidden"] = att.IsHidden;
                    row_att["AttributeIsManualDataEntry"] = att.IsManualDataEntry;
                    row_att["AttributeTrait"] = att.Trait;
                    row_att["AttributeIsConfigurationItem"] = att.IsConfigurationItem;
                    row_att["AttributeIsExcluded"] = att.IsExcluded;
                    row_att["AttributeIsIndexed"] = att.IsIndexed;
                    row_att["AttributeTypeQualifier"] = att.TypeQualifier;
                    row_att["AttributeDefaultValue"] = att.GetValue(null);
                    row_att["AttributeDataReference"] = att.DataReferencePlugIn;
                    row_att["AttributeDisplayDigits"] = att.DisplayDigits;
                    row_att["AttributeType"] = att.Type.Name;
                    dt.Rows.Add(row_att);

                    foreach(AFAttributeTemplate att_child in att.AttributeTemplates)
                    {
                        DataRow row_att_child = dt.NewRow();
                        row_att_child["Selected(x)"] = "x";
                        row_att_child["Parent"] = afeltemplate.Name;
                        row_att_child["Name"] = att_child.Parent.Name+"|"+ att_child.Name;
                        row_att_child["ObjectType"] = att_child.Identity;
                        row_att_child["Categories"] = att_child.CategoriesString;
                        row_att_child["Description"] = att_child.Description;
/*                      row_att_child["BaseTemplate"] = "";
                        row_att_child["Type"] = "";
                        row_att_child["AllowElementToExtend"] = "";
                        row_att_child["BaseTemplateOnly"] = "";
                        row_att_child["NamingPattern"] = "";*/
                        row_att_child["AttributeConfigString"] = att_child.ConfigString;
                        row_att_child["AttributeDefaultUOM"] = att_child.DisplayUOM;
                        row_att_child["AttributeIsHidden"] = att_child.IsHidden;
                        row_att_child["AttributeIsManualDataEntry"] = att_child.IsManualDataEntry;
                        row_att_child["AttributeTrait"] = att_child.Trait;
                        row_att_child["AttributeIsConfigurationItem"] = att_child.IsConfigurationItem;
                        row_att_child["AttributeIsExcluded"] = att_child.IsExcluded;
                        row_att_child["AttributeIsIndexed"] = att_child.IsIndexed;
                        row_att_child["AttributeTypeQualifier"] = att_child.TypeQualifier;
                        row_att_child["AttributeDefaultValue"] = att_child.GetValue(null);
                        row_att_child["AttributeDataReference"] = att_child.DataReferencePlugIn;
                        row_att_child["AttributeDisplayDigits"] = att_child.DisplayDigits;
                        row_att_child["AttributeType"] = att_child.Type.Name;
                        dt.Rows.Add(row_att_child);
                    }
                }
                foreach (AFAnalysisTemplate afanal in afeltemplate.GetAllAnalysisTemplates())
                {                    
                    DataRow row_analysis = dt.NewRow();
                    row_analysis["Selected(x)"] = "x";
                    row_analysis["Parent"] = afeltemplate.Name;
                    row_analysis["Name"] = afanal.Name;
                    row_analysis["ObjectType"] = afanal.Identity;
                    row_analysis["Categories"] = afanal.CategoriesString;
                    row_analysis["Description"] = afanal.Description;
/*                  row_analysis["BaseTemplate"] = "";
                    row_analysis["Type"] = "";
                    row_analysis["AllowElementToExtend"] = "";
                    row_analysis["BaseTemplateOnly"] = "";
                    row_analysis["NamingPattern"] = "";
                    row_analysis["AttributeConfigString"] = "";
                    row_analysis["AttributeDefaultUOM"] = "";
                    row_analysis["AttributeIsHidden"] = "";
                    row_analysis["AttributeIsManualDataEntry"] = "";
                    row_analysis["AttributeTrait"] = "";
                    row_analysis["AttributeIsConfigurationItem"] = "";
                    row_analysis["AttributeIsExcluded"] = "";
                    row_analysis["AttributeIsIndexed"] = "";
                    row_analysis["AttributeTypeQualifier"] = "";
                    row_analysis["AttributeDefaultValue"] = "";
                    row_analysis["AttributeDataReference"] = "";
                    row_analysis["AttributeDisplayDigits"] = "";
                    row_analysis["AttributeType"] = "";*/
                    row_analysis["AnalysisRule"] = afanal.AnalysisRule.Name;
                    row_analysis["AnalysisRuleConfigString"] = afanal.AnalysisRule.SimplifiedConfigString;
                    row_analysis["AnalysisRuleVariableMapping"] = afanal.AnalysisRule.VariableMapping;
                    row_analysis["TimeRule"] = afanal.TimeRule.Name;
                    row_analysis["TimeRuleconfigString"] = afanal.TimeRule.ConfigString;
                    dt.Rows.Add(row_analysis);

                    foreach (AFAnalysisRule analRule in afanal.AnalysisRule.AnalysisRules)
                    {
/*                        Console.ForegroundColor = ConsoleColor.Magenta;
                        Console.Write(analRule.Parent.Name);*/
                       
                            DataRow row_analysisRule = dt.NewRow();
                            row_analysisRule["Selected(x)"] = "x";
                            row_analysisRule["Parent"] = afeltemplate.Name + "\\"+afanal.Name;
                            row_analysisRule["Name"] = "[1]";
                            row_analysisRule["ObjectType"] = "TemplateAnalysisRule";
                            row_analysisRule["AnalysisRule"] = analRule.PlugIn;
                            row_analysisRule["AnalysisRuleConfigString"] = analRule.ConfigString;
                        dt.Rows.Add(row_analysisRule);
                    }
                }
                
            }
            return dt;
        }
        public static void GetAttributeTemplate(AFElementTemplate afeltemp)
        {
            AFNamedCollection<AFAttributeTemplate> attTemplList = afeltemp.AttributeTemplates;
            foreach (AFAttributeTemplate attTemp in attTemplList)
            {
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine(attTemp.Name);
            }
        }
        public static AFElement GetKitAFElement(AFDatabase afdb, string unitkit)
        {
            AFNamedCollectionList<AFElement> AFElementSearch = AFElement.FindElements(afdb, null, unitkit, OSIsoft.AF.AFSearchField.Name, true, OSIsoft.AF.AFSortField.Name, OSIsoft.AF.AFSortOrder.Ascending, 10000);
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(AFElementSearch[0].Name);
            return AFElementSearch[0];
        }
        public static void GetKitAFElement_parent(AFElement afel)
        {
            AFElement afel_parent = afel.Parent;
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine(afel_parent.Name);
        }
        public static void GetChildrenElementFromUnitKitElement(AFDatabase afdb, AFElement afel)
        {
            AFNamedCollectionList<AFElement> AFElementSearch = AFElement.FindElements(afdb, afel, null, OSIsoft.AF.AFSearchField.Name, true, OSIsoft.AF.AFSortField.Name, OSIsoft.AF.AFSortOrder.Ascending, 10000);
            foreach (AFElement el in AFElementSearch)
            {
                Console.ForegroundColor = ConsoleColor.Magenta;
                Console.WriteLine(el.Name);
                GetElementAttribute(el);
            }
        }
        public static void GetElementAttribute(AFElement afel)
        {
            AFNamedCollection<AFAttribute> attTemplList = afel.Attributes;
            foreach (AFAttribute att in attTemplList)
            {
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine(att.Name);
            }
        }
        public static List<AFElementTemplate> GetEFTemplateList(AFDatabase mydB)
        {
            AFNamedCollectionList<AFElementTemplate> aFEventFrameTemplates = new AFNamedCollectionList<AFElementTemplate>();
            aFEventFrameTemplates = AFElementTemplate.FindElementTemplates(mydB, "*GTP.Recommendation*", AFSearchField.Name, AFSortField.Name, AFSortOrder.Descending, 500);

           /* foreach (AFElementTemplate aFEventFrameTemplate in aFEventFrameTemplates)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("Event Frame template name: " + aFEventFrameTemplate.Name);
                GetAttributeTemplate(aFEventFrameTemplate);
            }*/
            return aFEventFrameTemplates.ToList();
        }
        public static DataTable EventFrameTemplateDataTable(List<AFElementTemplate> eftempList)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Selected(x)", typeof(string));
            dt.Columns.Add("Parent", typeof(string));
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("ObjectType", typeof(string));
            dt.Columns.Add("Description", typeof(string));
            dt.Columns.Add("BaseTemplate", typeof(string));
            dt.Columns.Add("AllowElementToExtend", typeof(string));
            dt.Columns.Add("BaseTemplateOnly", typeof(string));
            dt.Columns.Add("NamingPattern", typeof(string));
            dt.Columns.Add("Categories", typeof(string));
            dt.Columns.Add("CreationDate", typeof(string));
            dt.Columns.Add("ModifyDate", typeof(string));
            dt.Columns.Add("Severity", typeof(string));
            dt.Columns.Add("CanBeAckowledge", typeof(string));
            dt.Columns.Add("AttributeIsHidden", typeof(string));
            dt.Columns.Add("AttributeIsManualDataEntry", typeof(string));
            dt.Columns.Add("AttributeTrait", typeof(string));
            dt.Columns.Add("AttributeIsConfigurationItem", typeof(string));
            dt.Columns.Add("AttributeIsExcluded", typeof(string));
            dt.Columns.Add("AttributeIsIndexed", typeof(string));
            dt.Columns.Add("AttributeDefaultUOM", typeof(string));
            dt.Columns.Add("AttributeType", typeof(string));
            dt.Columns.Add("AttributeDefaultValue", typeof(string));
            dt.Columns.Add("AttributeDataReference", typeof(string));
            dt.Columns.Add("AttributeConfigString", typeof(string));
            dt.Columns.Add("AttributeDisplayDigits", typeof(string));

            DataRow row_init = dt.NewRow();
            row_init["Selected(x)"] = "Selected(x)";
            row_init["Parent"] = "Parent";
            row_init["Name"] = "Name";
            row_init["ObjectType"] = "ObjectType";
            row_init["Description"] = "Description";
            row_init["BaseTemplate"] = "BaseTemplate";
            row_init["AllowElementToExtend"] = "AllowElementToExtend";
            row_init["BaseTemplateOnly"] = "BaseTemplateOnly";
            row_init["NamingPattern"] = "NamingPattern";
            row_init["Categories"] = "Categories";
            row_init["CreationDate"] = "CreationDate";
            row_init["ModifyDate"] = "ModifyDate";
            row_init["Severity"] = "Severity";
            row_init["CanBeAckowledge"] = "CanBeAckowledge";
            row_init["AttributeIsHidden"] = "AttributeIsHidden";
            row_init["AttributeIsManualDataEntry"] = "AttributeIsManualDataEntry";
            row_init["AttributeTrait"] = "AttributeTrait";
            row_init["AttributeIsConfigurationItem"] = "AttributeIsConfigurationItem";
            row_init["AttributeIsExcluded"] = "AttributeIsExcluded";
            row_init["AttributeIsIndexed"] = "AttributeIsIndexed";
            row_init["AttributeDefaultUOM"] = "AttributeDefaultUOM";
            row_init["AttributeType"] = "AttributeType";
            row_init["AttributeDefaultValue"] = "AttributeDefaultValue";
            row_init["AttributeDataReference"] = "AttributeDataReference";
            row_init["AttributeConfigString"] = "AttributeConfigString";
            row_init["AttributeDisplayDigits"] = "AttributeDisplayDigits";
            dt.Rows.Add(row_init);

            foreach (AFElementTemplate eftemp in eftempList)
            {
                DataRow row = dt.NewRow();
                row["Selected(x)"] = "x";
                row["Parent"] = "";
                row["Name"] = eftemp.Name;
                row["ObjectType"] = "EventFrameTemplate";
                row["Description"] = eftemp.Description;
                row["BaseTemplate"] = eftemp.BaseTemplate != null ? eftemp.BaseTemplate.Name : "";
                row["AllowElementToExtend"] = eftemp.AllowElementToExtend;
                row["BaseTemplateOnly"] = eftemp.BaseTemplateOnly;
                row["NamingPattern"] = eftemp.NamingPattern;
                row["Categories"] = eftemp.CategoriesString;
                row["CreationDate"] = ""; //is it relevant?
                row["ModifyDate"] = "";
                row["Severity"] = eftemp.Severity;
                row["CanBeAckowledge"] = eftemp.CanBeAcknowledged;
/*                row["AttributeIsHidden"] = "";
                row["AttributeIsManualDataEntry"] = "";
                row["AttributeTrait"] = "";
                row["AttributeIsConfigurationItem"] = "";
                row["AttributeIsExcluded"] = "";
                row["AttributeIsIndexed"] = "";
                row["AttributeDefaultUOM"] = "";
                row["AttributeType"] = "";
                row["AttributeDefaultValue"] = "";
                row["AttributeDataReference"] = "";
                row["AttributeConfigString"] = "";
                row["AttributeDisplayDigits"] = "";*/
                dt.Rows.Add(row);

                foreach (AFAttributeTemplate att in eftemp.AttributeTemplates)
                {
                    DataRow row_att = dt.NewRow();
                    row_att["Selected(x)"] = "x";
                    row_att["Parent"] = eftemp.Name;
                    row_att["Name"] = att.Name;
                    row_att["ObjectType"] = att.Identity;
                    row_att["Categories"] = att.CategoriesString;
                    row_att["Description"] = att.Description;
                    row_att["AttributeIsHidden"] = att.IsHidden;
                    row_att["AttributeIsManualDataEntry"] = att.IsManualDataEntry;
                    row_att["AttributeTrait"] = att.Trait;
                    row_att["AttributeIsConfigurationItem"] = att.IsConfigurationItem;
                    row_att["AttributeIsExcluded"] = att.IsExcluded;
                    row_att["AttributeIsIndexed"] = att.IsIndexed;
                    row_att["AttributeDefaultUOM"] = att.DefaultUOM;
                    row_att["AttributeType"] = att.Type.Name;
                    row_att["AttributeDefaultValue"] = att.GetValue(null);
                    row_att["AttributeDataReference"] = att.DataReferencePlugIn;
                    row_att["AttributeConfigString"] = att.ConfigString ;
                    row_att["AttributeDisplayDigits"] = att.DisplayDigits;
                    dt.Rows.Add(row_att);
                }
                }
                return dt;
        }
    }
}
