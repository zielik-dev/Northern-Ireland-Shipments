using Microsoft.EntityFrameworkCore;
using Northern_Ireland_Shipments.Models.Commands;
using Northern_Ireland_Shipments.RemoteConfiguration.EcoSystemServerConfig;

namespace Northern_Ireland_Shipments.DbContext
{
    public class ApplicationLogsDbContext : Microsoft.EntityFrameworkCore.DbContext
    {
        private static string _DbConnectionLogsProd, _DbConnectionLogsTest, _environment;

        public ApplicationLogsDbContext(string environment)
        {
            _DbConnectionLogsProd = DbConnectionLogsProd.Read();
            _DbConnectionLogsTest = DbConnectionLogsTest.Read();
            _environment = environment;
        }
        public DbSet<MainLogsModel> MainLogsTable { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            if (_environment == "Production")
                optionsBuilder.UseSqlServer(_DbConnectionLogsProd);
            else
                optionsBuilder.UseSqlServer(_DbConnectionLogsTest);
        }
    }
}
