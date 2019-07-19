namespace DocxMould
{
    public class ReplaceStyle
    {
        public string FontFamily { get; set; }

        public int? FontSize { get; set; }

        public bool? Bold { get; set; }

        public bool? Capitalized { get; set; }

        public bool? DoubleStrikeThrough { get; set; }

        public bool? Embossed { get; set; }

        public bool? Imprinted { get; set; }

        public bool? Italic { get; set; }

        public bool? Shadowed { get; set; }

        public bool? SmallCaps { get; set; }

        public bool? StrikeThrough { get; set; }

        public VerticalPosition? VerticalPosition { get; set; }
    }

    public enum VerticalPosition
    {
        Baseline = 0,
        Superscript = 1,
        Subscript = 2
    }
}