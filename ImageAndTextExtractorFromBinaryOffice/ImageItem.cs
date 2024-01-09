using System;
using System.Collections.Generic;
using System.Text;

namespace ImageAndTextExtractorFromBinaryOffice
{
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
