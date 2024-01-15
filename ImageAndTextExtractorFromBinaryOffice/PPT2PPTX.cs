//using b2xtranslator.OfficeDrawing;
//using b2xtranslator.OpenXmlLib.PresentationML;
//using b2xtranslator.PptFileFormat;
//using b2xtranslator.PresentationMLMapping;
//using b2xtranslator.Shell;
//using b2xtranslator.StructuredStorage.Common;
//using b2xtranslator.StructuredStorage.Reader;
//using b2xtranslator.Tools;
//using System;
//using System.Globalization;
//using System.IO;
//using System.Threading;

//namespace ImageAndTextExtractorFromBinaryOffice
//{
//    public class ProcessingFile
//    {
//        public FileInfo File;

//        public ProcessingFile(string inputFile)
//        {
//            var inFile = new FileInfo(inputFile);

//            this.File = inFile.CopyTo(System.IO.Path.GetTempFileName(), true);
//            this.File.IsReadOnly = false;
//        }
//    }
//    public static class PPT2PPTX
//    {
//        public static bool ExtractFromPPT(string path)
//        {
//            return false;
//            // copy processing file
//            //var procFile = new ProcessingFile(path);

//            ////make output file name
//            //if (ChoosenOutputFile == null)
//            //{
//            //    if (InputFile.Contains("."))
//            //    {
//            //        ChoosenOutputFile = InputFile.Remove(InputFile.LastIndexOf(".")) + ".pptx";
//            //    }
//            //    else
//            //    {
//            //        ChoosenOutputFile = InputFile + ".pptx";
//            //    }
//            //}

//            ////
//            ////open the reader
//            //using (var reader = new StructuredStorageReader(procFile.File.FullName))
//            //{
//            //    // parse the ppt document
//            //    var ppt = new PowerpointDocument(reader);

//            //    // detect document type and name
//            //    var outType = b2xtranslator.PresentationMLMapping.Converter.DetectOutputType(ppt);
//            //    string conformOutputFile = Converter.GetConformFilename(ChoosenOutputFile, outType);

//            //    // create the pptx document
//            //    var pptx = PresentationDocument.Create(conformOutputFile, outType);

//            //    //start time
//            //    var start = DateTime.Now;
//            //    //TraceLogger.Info("Converting file {0} into {1}", InputFile, conformOutputFile);

//            //    // convert
//            //    Converter.Convert(ppt, pptx);

//            //    // stop time
//            //    var end = DateTime.Now;
//            //    var diff = end.Subtract(start);
//            //    TraceLogger.Info("Conversion of file {0} finished in {1} seconds", InputFile, diff.TotalSeconds.ToString(CultureInfo.InvariantCulture));
//            //}
//            ////
//            //var converter = new ConverterToXml.Converters.DocToDocx();
//            ////string curDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
//            ////string path = @"C:\01\test.doc"; // curDir +
//            //string curDir = Path.GetDirectoryName(path);
//            //converter.ConvertFromFileToDocxFile(path, curDir + @"\Files\Result.docx");

//            //if (!File.Exists(curDir + @"/Files/Result.docx"))
//            //{
//            //    return false;
//            //}

//            //ExtractResult result = new ExtractResult();
//            //System.Drawing.Image image;
//            //System.IO.Compression.ZipArchive documentArchive;
//            //string mediaPath = "word/media/";

//            ////string libPath = "libre/something/.../pictures/";
//            //// Open the document archive and loop through entries
//            //try
//            //{
//            //    using (documentArchive = new System.IO.Compression.ZipArchive(new FileStream(curDir + @"/Files/Result.docx", FileMode.Open)))
//            //    {
//            //        foreach (System.IO.Compression.ZipArchiveEntry entry in documentArchive.Entries)
//            //        {
//            //            // Check the location and if media try opening the image
//            //            if (entry.FullName.StartsWith(mediaPath))
//            //            {
//            //                try
//            //                {
//            //                    image = System.Drawing.Image.FromStream(entry.Open());
//            //                    result.ImageItems.Add(
//            //                       new ImageItem()
//            //                       {
//            //                           Name = entry.Name,
//            //                           Image = image
//            //                       });
//            //                }
//            //                catch (System.Exception exception)
//            //                {
//            //                    // Add information about the error
//            //                    result.Errors.Add(
//            //                       new ExtractError()
//            //                       {
//            //                           File = entry.Name,
//            //                           Error = exception.Message
//            //                       });
//            //                }
//            //            }
//            //        }
//            //    }
//            //}
//            //catch (System.Exception exception)
//            //{
//            //    //Assert.True(false);
//            //}
//            //string currentDirectory = Path.GetDirectoryName(path) + @"\Files\";
//            //foreach (var imageItem in result.ImageItems)
//            //{

//            //    string imagePath = Path.Combine(currentDirectory, imageItem.Name);

//            //    imageItem.Image.Save(imagePath);

//            //}
//            //return true;
//            ////Assert.True(File.Exists(curDir + @" / Files/Result.docx"));
//        }
//    }
//}
