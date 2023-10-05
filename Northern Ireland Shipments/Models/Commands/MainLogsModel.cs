using Microsoft.EntityFrameworkCore.Metadata.Internal;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel;

namespace Northern_Ireland_Shipments.Models.Commands
{
    public class MainLogsModel
    {
        [Key]
        public int LogId { get; set; }

        [Required]
        [Column(TypeName = "varchar(10)")]
        [DisplayName("Log Date")]
        public string LogDate { get; set; }

        [Required]
        [Column(TypeName = "varchar(8)")]
        [DisplayName("Log Time")]
        public string LogTime { get; set; }

        [Required]
        [Column(TypeName = "varchar(250)")]
        [DisplayName("Application Name")]
        public string ApplicationName { get; set; }

        [Required]
        [Column(TypeName = "varchar(250)")]
        [DisplayName("Process Name")]
        public string ProcessName { get; set; }

        [Required]
        [Column(TypeName = "varchar(250)")]
        [DisplayName("Status Name")]
        public string Status { get; set; }
    }
}