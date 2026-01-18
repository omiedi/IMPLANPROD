namespace IMPLANPROD.Server.Data.Interfaces
{
    public interface IAuditable
    {
        DateTime AddRecord { get; set; }
        DateTime LastUpdate { get; set; }
    }
}
