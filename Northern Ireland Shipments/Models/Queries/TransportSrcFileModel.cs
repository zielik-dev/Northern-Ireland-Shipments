namespace Northern_Ireland_Shipments.Models.Queries
{
    public class TransportSrcFileModel
    {
        public DateTime Date { get; set; }
        public string? Client_Job_Number { get; set; }
        public string? Document_Reference { get; set; }
        public string? Header_Information { get; set; }
        public string? Pallet_Count { get; set; }
        public string? Vehicle_Reg { get; set; }
        public string? Trailer_Org { get; set; }
        public string? Trailer_Sys { get; set; }
        public string? Gmr { get; set; }
        public string? Completed_By { get; set; }
        public string? Comment { get; set; }
    }
}