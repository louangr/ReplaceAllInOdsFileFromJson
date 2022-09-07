namespace Utils
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using Aspose.Cells;
    using Newtonsoft.Json;
    
    public static class Process
    {
        #region Privates fields

        private static readonly string ODS_SOURCE_FILE_PATH = "source.ods";
        private static readonly string JSON_SOURCE_FILE_PATH = "source.json";
        private static readonly string ODS_OUTPOUT_FILE_PATH = "output.ods";

        private static List<Person> persons;
        private static Workbook workbook;

        #endregion Privates fields

        #region Publics methods

        public static void ReplaceAll(ReplaceWay way)
        {
            LoadPersonsAndOdsSource();

            if (persons != null && persons.Count > 0 && workbook != null)
            {
                SearchAndReplace(way);
                workbook.Save(ODS_OUTPOUT_FILE_PATH);
            }
        }
        
        #endregion Publics methods

        #region Privates methods
        
        private static void SearchAndReplace(ReplaceWay way)
        {
            foreach (Person p in persons)
            {
                var oldValue = way == ReplaceWay.IdToName ? p.Id.ToString() : p.Name;
                var newValue = way == ReplaceWay.IdToName ? p.Name : p.Id.ToString();
                
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    FindOptions opts = new FindOptions();
                    opts.LookInType = LookInType.Values;
                    opts.LookAtType = LookAtType.Contains;
                    opts.RegexKey = true;
                    Cell cell = null;
                    do
                    {
                        cell = sheet.Cells.Find(oldValue, cell, opts);
                        if (cell != null)
                        {
                            string celltext = cell.Value.ToString();
                            celltext = celltext.Replace(oldValue, newValue);
                            cell.PutValue(celltext);
                        }
                    }
                    while (cell != null);
                }
            }
        }

        private static void LoadPersonsAndOdsSource()
        {
            try
            {
                if (File.Exists(JSON_SOURCE_FILE_PATH))
                {
                    using (var stream = File.OpenRead(JSON_SOURCE_FILE_PATH))
                    using (var sr = new StreamReader(stream))
                    using (var jsonTextReader = new JsonTextReader(sr))
                    {
                        
                        persons = new JsonSerializer().Deserialize<List<Person>>(jsonTextReader);
                    }
                }

                workbook = new Workbook(ODS_SOURCE_FILE_PATH);
            }
            catch (Exception)
            {
            }
        }
        
        #endregion Privates methods
    }
}
