using ImageAndTextExtractorFromBinaryOffice.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ImageAndTextExtractorFromBinaryOffice
{
    public static class FileManager
    {
        /// <summary>
        /// Возвращает путь к документу из аргументов на этапе открытия нашего прокси.
        /// </summary>
        /// <param name="e">Событие при открытии прокси.</param>
        /// <returns>Путь к документу.</returns>
        public static bool FileIsExistFromStartUpArguments(string[] args,out string filePath)
        {
            for (int i = 0; i != args.Length; i++)
            {
                if (File.Exists(args[i]))
                {
                    //DSSLogger.Trace($"FullFilePathFromStartUpArguments: {e.Args[i]}.");
                    filePath = args[i];
                    return true;
                }
            }
            filePath = string.Empty;
            return false;
        }
        public static bool GetOpenFileCommandByExtension(string extension, out Extension supportedExtension)
        {
            if (String.IsNullOrEmpty(extension))
            {
                supportedExtension = new OtherExtension();
                return false;
            }
            switch (extension)
            {
                case ".docx":
                    {
                        supportedExtension = new DocxExtension();
                        return true;
                    }
                case ".doc":
                    {
                        supportedExtension = new DocExtension();
                        return true;
                    }
                case ".odt":
                    {
                        supportedExtension = new OdtExtension();
                        return true;
                    }
                case ".xls":
                    {
                        supportedExtension = new XlsExtension();
                        return true;
                    }
                case ".ppt":
                    {
                        supportedExtension = new OtherExtension();
                        return true;
                    }
                case ".vsd":
                    {
                        supportedExtension = new OtherExtension();
                        return true;
                    }
                default:
                    {
                        supportedExtension = new OtherExtension();
                        return false;
                    }
            }
        }
    }
}
