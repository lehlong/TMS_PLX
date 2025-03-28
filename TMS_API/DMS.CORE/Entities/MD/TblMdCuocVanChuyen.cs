using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using DMS.CORE.Common;

namespace DMS.CORE.Entities.MD
{
    [Table("T_MD_CUOC_VAN_CHUYEN")]
    public class TblMdCuocVanChuyen : BaseEntity
    {
        [Key]
        [Column("CODE", TypeName = "VARCHAR(50)")]
        public string? Code { set; get; }

        [Column("HEADER_CODE", TypeName = "VARCHAR(50)")]
        public string? HeaderCode { get; set; }

        [Column("VSART", TypeName = "VARCHAR(50)")]
        public string? VSART { get; set; }

        [Column("TDLNR", TypeName = "VARCHAR(50)")]
        public string? TDLNR { get; set; }

        [Column("KNOTA", TypeName = "VARCHAR(50)")]
        public string? KNOTA { get; set; }

        [Column("OIGKNOTD", TypeName = "VARCHAR(50)")]
        public string? OIGKNOTD { get; set; }

        [Column("MATNR", TypeName = "VARCHAR(50)")]
        public string? MATNR { get; set; }

        [Column("VRKME", TypeName = "VARCHAR(50)")]
        public string? VRKME { get; set; }

        [Column("KBETR", TypeName = "DECIMAL(18, 0)")]
        public decimal? KBETR { get; set; }

        [Column("KONWA", TypeName = "VARCHAR(50)")]
        public string? KONWA { get; set; }

        [Column("KPEIN", TypeName = "VARCHAR(50)")]
        public string? KPEIN { get; set; }

        [Column("KMEIN", TypeName = "VARCHAR(50)")]
        public string? KMEIN { get; set; }

        [Column("DATAB", TypeName = "VARCHAR(50)")]
        public string? DATAB { get; set; }

        [Column("DATBI", TypeName = "VARCHAR(50)")]
        public string? DATBI { get; set; }

        [Column("ROW_INDEX", TypeName = "INT")]
        public int? RowIndex { get; set; }
    }

}
