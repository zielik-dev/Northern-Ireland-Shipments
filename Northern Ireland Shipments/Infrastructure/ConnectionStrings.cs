namespace Northern_Ireland_Shipments.Infrastructure
{
    public class ConnectionStrings
    {
        public static string ecoSystemGeneralConfig = @"\\XXX\LocalData$\XXX\.NET Core Solution - Production Env\.Config\EcoSystemGeneralConfig.xml";
        public static string ecoSystemDbConnection = @"\\XXX\LocalData$\XXX\.NET Core Solution - Production Env\.Config\EcoSystemDatabaseConfig.xml";
        public static string ecoSystemEmailDistributionConfig = @"\\XXX\LocalData$\XXX\.NET Core Solution - Production Env\.Config\EcoSystemEmailDistributionConfig.xml";
        public static string smtpConfigFile = @"\\XXX\LocalData$\XXX\.NET Core Solution - Production Env\.Config\Smtp.server.config.xml";

        public static string reportTemplate = @"\\XXX\LocalData$\XXX\.NET Core Solution - Production Env\.Templates\Northern Ireland Shipments Report.xlsm";
        public static string sheetTemplate = "Report";

        public static string archiveEndDir = "Northern Ireland Shipments";

        public static string queryFile = "Northern_Ireland_Shipments_Query.sql";

        public static string sourceWb = @"\\XXX\LocalData$\XXX\TRANSPORT(Restored)\Northern Ireland\Northern Ireland Shipments.xlsm";
        public static string sheetSrc = "Summary";

        public static string logTitle = "Northern Ireland Shipments";
    }
}