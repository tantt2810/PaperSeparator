namespace PaperSeparator.DataTransfer
{
    public class Configure
    {
        public int Column { get; set; }
        public string PageSetupOrientation { get; set; }
        public string PageSetupPaperSize { get; set; }
        public int PageSetupLeftMargin { get; set; }
        public int PageSetupRightMargin { get; set; }
        public int PageSetupTopMargin { get; set; }
        public int PageSetupBottomMargin { get; set; }
        public int CellFormatTopPadding { get; set; }
        public int CellFormatBottomPadding { get; set; }
        public int CellFormatLeftPadding { get; set; }
        public int CellFormatRightPadding { get; set; }
        public int RowFormatHeight { get; set; }
        public string RowFormatHeightRule { get; set; }
        public string ParagraphFormatAlignment { get; set; }
        public string CellFormatVerticalAlignment { get; set; }
        public int FontSize { get; set; }
        public string FontName { get; set; }
        public string FontBold { get; set; }
        public string TableAllowAutoFit { get; set; }
        public string DataPath { get; set; }
        public string SavePath { get; set; }
        public string FileName { get; set; }
    }
}
