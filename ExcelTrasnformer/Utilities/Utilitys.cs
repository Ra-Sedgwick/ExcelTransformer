using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTrasnformer.Utilities
{
    public static class Utilitys
    {

        public static void ImportFile(String filePath)
        {
            FileStream fileStream = null;

            try
            {
                fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using (TextReader tr = new StreamReader(fileStream))
                {
                    fileStream = null;
                    // Code here
                }
            }
            finally
            {
                if (fileStream != null)
                    fileStream.Dispose();
            }

        }

    }

}
