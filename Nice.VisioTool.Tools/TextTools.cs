using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Nice.VisioTool.Tools
{
    public class TextTools
    {
        private string FilePath;
        private string FileName;        
        public TextTools(string filePath, string fileName)
        {
            FilePath = filePath;
            FileName = fileName;
        }
/// <summary>
/// Method for creating file in case it does not exist
/// </summary>
        public void CreateFile()
        {
            string FullPath = string.Empty;
            try
            {
                FullPath = string.Format("{0}\\{1}", FilePath, FileName);
                if (!File.Exists(FullPath))
                {
                    File.Create(FullPath).Close();
                }
            }
            catch (Exception ex)
            {
            }            
        }
        public void WriteToFile(string contents)
        {
            string FullPath = string.Empty;
            try
            {
                FullPath = string.Format("{0}\\{1}", FilePath, FileName);
                if (File.Exists(FullPath))
                {
                    File.WriteAllText(FullPath,contents);
                }
            }
            catch (Exception ex)
            {
            }
        }

    }
}
