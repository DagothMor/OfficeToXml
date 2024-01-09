using System;
using System.Collections.Generic;
using System.Text;

namespace ImageAndTextExtractorFromBinaryOffice
{
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
}
