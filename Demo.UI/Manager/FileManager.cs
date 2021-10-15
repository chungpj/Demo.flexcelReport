using System;
using System.Configuration;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;

namespace Demo.UI.Manger
{
    public class FileManager
    {
        public FileManager()
        {
            var filedata = ConfigurationManager.AppSettings["FileData"];
            var appdata = ConfigurationManager.AppSettings["AppData"];
            var fileFolder = ConfigurationManager.AppSettings["FileFolder"];
            var temFolder = ConfigurationManager.AppSettings["TempFolder"];
            var cacheFolder = ConfigurationManager.AppSettings["CacheFolder"];
            var indexFolder = ConfigurationManager.AppSettings["IndexFolder"];
            var scanFolder = ConfigurationManager.AppSettings["ScanFolder"];

            var locationPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase.Replace(@"file:///", "")), "..");
            FileFolder = Path.Combine(filedata, fileFolder + Path.DirectorySeparatorChar);
            TemFolder = Path.Combine(locationPath, appdata, temFolder + Path.DirectorySeparatorChar);
            CacheFolder = Path.Combine(filedata, cacheFolder + Path.DirectorySeparatorChar);
            IndexFolder = Path.Combine(locationPath, appdata, indexFolder + Path.DirectorySeparatorChar);
            ScanFolder = Path.Combine(filedata, scanFolder + Path.DirectorySeparatorChar);
            FileData = Path.Combine(filedata);
        }
        public string FileFolder { get; private set; }
        public string TemFolder { get; private set; }
        public string CacheFolder { get; private set; }
        public string IndexFolder { get; private set; }
        public string ScanFolder { get; private set; }
        public string FileData { get; private set; }
        public string CreateFilePath(string fileName, DateTime? dateTime = null)
        {
            fileName = Replate(fileName);
            var date = dateTime ?? DateTime.Now;
            var folder = Path.Combine(FileFolder, date.Year.ToString(), date.Month.ToString(), date.Day.ToString());
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);
            var name = Guid.NewGuid() + fileName;
            return Path.Combine(date.Year.ToString(), date.Month.ToString(), date.Day.ToString(), name);
        }

        public string Replate(string value)
        {
            string regex = $"[{Regex.Escape(new string(Path.GetInvalidFileNameChars()))}]";
            Regex removeInvalidChars = new Regex(regex, RegexOptions.Singleline | RegexOptions.Compiled | RegexOptions.CultureInvariant);
            value = removeInvalidChars.Replace(value, "_");
            return value;
        }
        public string GetFilePath(string filePath)
        {
            return Path.Combine(FileFolder, filePath);
        }
        public string CreateCacheFilePath(string fileName)
        {
            var date = DateTime.Now;
            var folder = Path.Combine(CacheFolder, date.ToString("ddMMyyyy"));
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);
            var name = Guid.NewGuid() + fileName;
            return Path.Combine(folder, name);
        }
        public string GetTempFilePath(string filename)
        {
            return Path.Combine(TemFolder, filename);
        }
    }
}