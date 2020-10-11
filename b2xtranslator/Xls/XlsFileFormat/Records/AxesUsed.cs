﻿using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies the number of axis groups on the chart.
    /// </summary>
    [BiffRecord(RecordType.AxesUsed)]
    public class AxesUsed : BiffRecord
    {
        public const RecordType ID = RecordType.AxesUsed;

        public enum AxisGroupsPresent : ushort
        {
            /// <summary>
            /// A single primary axis group is present
            /// </summary>
            Primary = 0x1,

            /// <summary>
            /// Both a primary axis group and a secondary axis group are present
            /// </summary>
            PrimaryAndSecondary = 0x2
        }

        /// <summary>
        /// An unsigned integer that specifies the number of axis groups on the chart.
        /// </summary>
        public AxisGroupsPresent cAxes = AxisGroupsPresent.Primary;

        public AxesUsed(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.cAxes = (AxisGroupsPresent)reader.ReadUInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
