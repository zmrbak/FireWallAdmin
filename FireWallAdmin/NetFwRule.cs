using NetFwTypeLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FireWallAdmin
{
    public class NetFwRule
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public bool Enabled { get; set; } = true;
        public NET_FW_ACTION_ Action { get; set; } = NET_FW_ACTION_.NET_FW_ACTION_ALLOW;
        public NET_FW_RULE_DIRECTION_ Direction { get; set; } = NET_FW_RULE_DIRECTION_.NET_FW_RULE_DIR_IN;
        public string LocalAddresses { get; set; } = "*";
        public int Protocol { get; set; } = 6;
        public List<Ports> LocalPorts { get; set; }
        public string InterfaceTypes { get; set; } = "ALL";
        public List<RemoteAddress> RemoteAddresses { get; set; }
        public string RemotePorts { get; set; } = "*";
    }

    public class Ports
    {
        public int PortNumber { get; set; }
        public string Description { get; set; }
    }

    public class RemoteAddress
    {
        public string IpAddress { get; set; }
        public string Description { get; set; }
    }
}
