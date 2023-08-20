namespace ExcelManager
{
    public abstract class ExcelModifier : ExcelReader
    {
        public ExcelReader ExcelToCopyFrom { get; set; }
        public ExcelModifier(string fileElia, string filePathToCopyFrom) : base(fileElia)
        {
            ExcelToCopyFrom = new ExcelReader(filePathToCopyFrom);
        }
        public abstract bool Allign(string sheetNameFileElia, string sheetNameFileToCopyFrom);
    }
}
