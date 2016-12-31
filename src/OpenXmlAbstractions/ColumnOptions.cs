using System.Collections.Generic;

namespace OpenXmlAbstractions
{
    public class ColumnOptions
    {
        public bool WrapText { get; set; } = true;

        public uint TextRotation { get; set; } = 0;

        public string TextColor { get; set; } = "000";

        public IList<string> TextReplacements { get; set; }

        public Dictionary<string, string> TextColorChanges { get; set; }
    }
}