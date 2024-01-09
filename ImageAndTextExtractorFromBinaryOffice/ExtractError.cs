using System;
using System.Collections.Generic;
using System.Text;

namespace ImageAndTextExtractorFromBinaryOffice
{
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
}
