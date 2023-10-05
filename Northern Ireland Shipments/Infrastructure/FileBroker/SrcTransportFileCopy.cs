using Northern_Ireland_Shipments.Interfaces;
using Northern_Ireland_Shipments.RemoteConfiguration.EcoSystemServerConfig;

namespace Northern_Ireland_Shipments.Infrastructure.FileBroker
{
    public class SrcTransportFileCopy : ConnectionStrings, ISrcTransportFileCopy
    {
        private static SrcTransportFileCopy srcTransportFileCopy;
        private readonly string inboundDir;

        public SrcTransportFileCopy()
        {
            inboundDir = InboundDir.Read();
        }

        public static SrcTransportFileCopy Instance
        {
            get
            {
                if (srcTransportFileCopy == null)
                    srcTransportFileCopy = new SrcTransportFileCopy();
                return srcTransportFileCopy;
            }
        }

        public string CopySrcTransportFileToInbound()
        {
            string transportFileName = Path.GetFileName(sourceWb);
            string inboundFile = Path.Combine(inboundDir, transportFileName);

            if(File.Exists(inboundFile))
                File.Delete(inboundFile);

            if (File.Exists(sourceWb))
                File.Copy(sourceWb, inboundFile);

            return inboundFile;
        }
    }
}