namespace Northern_Ireland_Shipments.Interfaces
{
    public interface IArchiveTemplate
    {
        public void CopyTemplateToArchive(string environment, DateTime dt);
    }
}
