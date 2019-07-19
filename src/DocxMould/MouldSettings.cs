using System;
using System.Collections.Generic;
using System.Text;

namespace DocxMould
{
    public static class MouldSettings
    {
        public static string FieldPrefix { get; set; } = "{$";

        public static string FieldSuffix { get; set; } = "}";

        public static string SectionPrefix { get; set; } = "{Section$";

        public static string SectionSuffix { get; set; } = "$}";
    }
}
