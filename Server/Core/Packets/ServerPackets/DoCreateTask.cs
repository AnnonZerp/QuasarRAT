using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using xServer.Core.Networking;

namespace xServer.Core.Packets.ServerPackets
{
    [Serializable]
    public class DoCreateTask : IPacket
    {
        public string TaskName { get; set; }
        public string Path { get; set; }
        public string Arguments { get; set; }

        public DoCreateTask(string taskName, string path, string arguments = null)
        {
            TaskName = taskName;
            Path = path;
            Arguments = arguments;
        }

        public void Execute(Client client)
        {
            client.Send(this);
        }
    }
}
