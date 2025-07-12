using System.ComponentModel.DataAnnotations;

namespace ExcelDateImporter.Api.Models
{
    public class ExcelDateEntry
    {
        [Key]
        public int Id { get; set; }
        
        public string OriginalValue { get; set; } = string.Empty;
        
        public DateTime? ParsedDate { get; set; }
        
        public string DateFormat { get; set; } = string.Empty;
        
        public bool IsSuccessfullyParsed { get; set; }
        
        public string ErrorMessage { get; set; } = string.Empty;
        
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    }
}