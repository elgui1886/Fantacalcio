namespace ExcelManager
{
    public abstract class ExcelModifier : ExcelReader
    {
        public ExcelReader ExcelToCopyFrom { get; set; }
        public ExcelModifier(string filePath, ExcelReader excelToCopyFrom) : base(filePath)
        {
            ExcelToCopyFrom = excelToCopyFrom;
        }
        public abstract bool Allign(string sheetNameFileElia, string sheetNameFileToCopyFrom);
    }
}
