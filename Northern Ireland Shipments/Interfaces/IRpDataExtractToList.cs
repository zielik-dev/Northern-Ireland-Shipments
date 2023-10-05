using Northern_Ireland_Shipments.Models.Queries;

namespace Northern_Ireland_Shipments.Interfaces
{
    public interface IRpDataExtractToList
    {
        public List<RpDbQueryModel> GetList();
    }
}
