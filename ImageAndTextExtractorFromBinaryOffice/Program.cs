using ConverterToXml.Converters;
//using ConverterToXml.Test;
using ImageAndTextExtractorFromBinaryOffice.Extensions;
using System;
using System.IO;
using System.Xml;

namespace ImageAndTextExtractorFromBinaryOffice
{
    class Program
    {
        static void Main(string[] args)
        {
            // если в аргументах ничего нет то уведомляем, выходим
            if (args.Length == 0)
            {
                Console.WriteLine("args is null");
                Console.Read();
                return;
            }
            // если в аргументах нету пути к файлу то уведомляем и выходим
            if (!FileManager.FileIsExistFromStartUpArguments(args, out string filePath))
            {
                Console.WriteLine("File path not found.");
                Console.Read();
                return;
            }
            // если в пути нету поддерживаемого расширения, то уведомляем и выходим
            if (!FileManager.GetOpenFileCommandByExtension(Path.GetExtension(filePath), out Extension supportedExtension))
            {
                Console.WriteLine("Extension was not supported");
                Console.Read();
                return;
            }
            var workingPath = Path.GetDirectoryName(filePath) + @"\Files"; // Directory.GetCurrentDirectory()
            // Determine whether the directory exists.
            if (!Directory.Exists(workingPath))
            {
                DirectoryInfo di = Directory.CreateDirectory(workingPath);
                di.Create();
            }
            else 
            {
                DirectoryInfo di = Directory.CreateDirectory(workingPath);
                di.Delete(true);
                di.Create();
            }
            // начало обработки файла в зависимости от расширения, скорее всего внутри будет еще логика по понятию того кем был поражден файл, что там со структурой.

            //ImageExtractor(file);

            switch (supportedExtension)
            {
                case OtherExtension _:
                    // handle OtherExtension
                    break;
                case DocExtension _:
                    Console.WriteLine("Extracting from DOCX file...");
                    ExtractFromDOCX(filePath);
                    break;
                case PptExtension _:
                    Console.WriteLine("Extracting from ppt file...");
                    ExtractFromDOCX(filePath);
                    break;
                case OdtExtension _:
                    // handle OdtExtension
                    break;
                default:
                    // handle other cases
                    break;
            }
        }
        public static bool ExtractFromPPT(string path)
        {
            var converter = new ConverterToXml.Converters.DocToDocx();
            //string curDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            //string path = @"C:\01\test.doc"; // curDir +
            string curDir = Path.GetDirectoryName(path);
            converter.ConvertFromFileToDocxFile(path, curDir + @"\Files\Result.docx");

            if (!File.Exists(curDir + @"/Files/Result.docx"))
            {
                return false;
            }

            ExtractResult result = new ExtractResult();
            System.Drawing.Image image;
            System.IO.Compression.ZipArchive documentArchive;
            string mediaPath = "word/media/";

            //string libPath = "libre/something/.../pictures/";
            // Open the document archive and loop through entries
            try
            {
                using (documentArchive = new System.IO.Compression.ZipArchive(new FileStream(curDir + @"/Files/Result.docx", FileMode.Open)))
                {
                    foreach (System.IO.Compression.ZipArchiveEntry entry in documentArchive.Entries)
                    {
                        // Check the location and if media try opening the image
                        if (entry.FullName.StartsWith(mediaPath))
                        {
                            try
                            {
                                image = System.Drawing.Image.FromStream(entry.Open());
                                result.ImageItems.Add(
                                   new ImageItem()
                                   {
                                       Name = entry.Name,
                                       Image = image
                                   });
                            }
                            catch (System.Exception exception)
                            {
                                // Add information about the error
                                result.Errors.Add(
                                   new ExtractError()
                                   {
                                       File = entry.Name,
                                       Error = exception.Message
                                   });
                            }
                        }
                    }
                }
            }
            catch (System.Exception exception)
            {
                //Assert.True(false);
            }
            string currentDirectory = Path.GetDirectoryName(path) + @"\Files\";
            foreach (var imageItem in result.ImageItems)
            {

                string imagePath = Path.Combine(currentDirectory, imageItem.Name);

                imageItem.Image.Save(imagePath);

            }
            return true;
            //Assert.True(File.Exists(curDir + @" / Files/Result.docx"));
        }
        public static bool ExtractFromDOCX(string path)
        {
            ConverterToXml.Converters.DocToDocx converter = new ConverterToXml.Converters.DocToDocx();
            //string curDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            //string path = @"C:\01\test.doc"; // curDir +
            string curDir = Path.GetDirectoryName(path);
            converter.ConvertFromFileToDocxFile(path, curDir + @"\Files\Result.docx");

            if (!File.Exists(curDir + @"/Files/Result.docx"))
            {
                return false;
            }

            ExtractResult result = new ExtractResult();
            System.Drawing.Image image;
            System.IO.Compression.ZipArchive documentArchive;
            string mediaPath = "word/media/";

            //string libPath = "libre/something/.../pictures/";
            // Open the document archive and loop through entries
            try
            {
                using (documentArchive = new System.IO.Compression.ZipArchive(new FileStream(curDir + @"/Files/Result.docx", FileMode.Open)))
                {
                    foreach (System.IO.Compression.ZipArchiveEntry entry in documentArchive.Entries)
                    {
                        // Check the location and if media try opening the image
                        if (entry.FullName.StartsWith(mediaPath))
                        {
                            try
                            {
                                image = System.Drawing.Image.FromStream(entry.Open());
                                result.ImageItems.Add(
                                   new ImageItem()
                                   {
                                       Name = entry.Name,
                                       Image = image
                                   });
                            }
                            catch (System.Exception exception)
                            {
                                // Add information about the error
                                result.Errors.Add(
                                   new ExtractError()
                                   {
                                       File = entry.Name,
                                       Error = exception.Message
                                   });
                            }
                        }
                    }
                }
            }
            catch (System.Exception exception)
            {
                //Assert.True(false);
            }
            string currentDirectory = Path.GetDirectoryName(path) + @"\Files\";
            foreach (var imageItem in result.ImageItems)
            {

                string imagePath = Path.Combine(currentDirectory, imageItem.Name);

                imageItem.Image.Save(imagePath);

            }
            return true;
            //Assert.True(File.Exists(curDir + @" / Files/Result.docx"));
        }
    }
}

// ms -> doc -> docx - все нормально 
// 

