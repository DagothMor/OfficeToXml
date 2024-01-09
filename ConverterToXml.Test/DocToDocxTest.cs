using System.IO;
using Xunit;

namespace ConverterToXml.Test
{
    public class DocToDocxTest
    {
        [Fact]
        public void DocToDocxConvertToDocx_NotNull()
        {
            Converters.DocToDocx converter = new Converters.DocToDocx();
            string curDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string path =  @"C:\01\test.doc"; // curDir +
            converter.ConvertFromFileToDocxFile(path, curDir + @"/Files/Result.docx");

            if(!File.Exists(curDir + @"/Files/Result.docx"))
            {
                Assert.True(false);
            }

            ExtractResult result = new ExtractResult();
            System.Drawing.Image image;
            System.IO.Compression.ZipArchive documentArchive;
            string mediaPath = "word/media/";

            string libPath = "libre/something/.../pictures/";
            // Open the document archive and loop through entries
            try
            {
                using (documentArchive = new System.IO.Compression.ZipArchive(new FileStream(curDir + @"/Files/Result.docx",FileMode.Open)))
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
                Assert.True(false);
            }
            string currentDirectory = Directory.GetCurrentDirectory();
            foreach (var imageItem in result.ImageItems)
            {
                
                    string imagePath = Path.Combine(currentDirectory, imageItem.Name);

                    imageItem.Image.Save(imagePath);
                
            }
            Assert.True(File.Exists(curDir + @"/Files/Result.docx"));
        }
    }
    /// <summary>
    /// The result of the extraction
    /// </summary>
    public class ExtractResult
    {
        /// <summary>
        /// File name which was investigated
        /// </summary>
        public string File { get; set; }
        /// <summary>
        /// Image items found
        /// </summary>
        public System.Collections.Generic.List<ImageItem> ImageItems { get; private set; }
        /// <summary>
        /// Errors encountered
        /// </summary>
        public System.Collections.Generic.List<ExtractError> Errors { get; private set; }

        /// <summary>
        /// Default constructor
        /// </summary>
        public ExtractResult()
        {
            this.ImageItems = new System.Collections.Generic.List<ImageItem>();
            this.Errors = new System.Collections.Generic.List<ExtractError>();
        }
    }

    /// <summary>
    /// This class represents an error occurred during image extraction
    /// </summary>
    public class ExtractError
    {
        /// <summary>
        /// File name
        /// </summary>
        public string File { get; set; }
        /// <summary>
        /// Error description
        /// </summary>
        public string Error { get; set; }
    }

    /// <summary>
    /// This class represents an image found
    /// </summary>
    public class ImageItem
    {
        /// <summary>
        /// NAme of the image
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// The image
        /// </summary>
        public System.Drawing.Image Image { get; set; }
    }


}
