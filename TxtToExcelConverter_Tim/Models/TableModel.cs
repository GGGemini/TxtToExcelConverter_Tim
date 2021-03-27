namespace TxtToExcelConverter_Tim.Models
{
    public class TableModel
    {
        /// <summary>
        /// Column A
        /// </summary>
        public string ComponentType { get; set; } = "-";
        /// <summary>
        /// Column B
        /// </summary>
        public string Name { get; set; } = "-";
        /// <summary>
        /// Column C
        /// </summary>
        public string Nominal { get; set; } = "-";
        /// <summary>
        /// Column D
        /// </summary>
        public string Deviation { get; set; } = "-";
        /// <summary>
        /// Column E
        /// </summary>
        public string CaseType { get; set; } = "-";
        /// <summary>
        /// Column F
        /// </summary>
        public string Comment { get; set; } = "-";
        /// <summary>
        /// Columns G and H
        /// </summary>
        public QuanityModel Quanity { get; set; } = new QuanityModel();
        /// <summary>
        /// Column I
        /// </summary>
        public string Remark { get; set; }
    }
}