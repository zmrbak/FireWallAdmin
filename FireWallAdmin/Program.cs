using NetFwTypeLib;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace FireWallAdmin
{
    class Program
    {
        string fireWallConfigFile = "fw.json";
        static void Main(string[] args)
        {
            Program program = new Program();
            program.Run();
        }

        public void Run()
        {
            if (ConfigFileCheck() == false) return;
            string jsonRules = File.ReadAllText(fireWallConfigFile);
            List<NetFwRule> netFwRules = (List<NetFwRule>)(new JavaScriptSerializer().Deserialize(jsonRules, typeof(List<NetFwRule>)));

            //实例化COM引用对象
            Type type = Type.GetTypeFromProgID("HNetCfg.FwPolicy2");
            INetFwPolicy2 INetFwPolicy = (INetFwPolicy2)Activator.CreateInstance(type);

            //启用防火墙
            INetFwPolicy.FirewallEnabled[NET_FW_PROFILE_TYPE2_.NET_FW_PROFILE2_PUBLIC] = true;
            INetFwPolicy.FirewallEnabled[NET_FW_PROFILE_TYPE2_.NET_FW_PROFILE2_PRIVATE] = true;
            INetFwPolicy.FirewallEnabled[NET_FW_PROFILE_TYPE2_.NET_FW_PROFILE2_DOMAIN] = true;

            //添加策略
            foreach (NetFwRule item in netFwRules)
            {
                //删除旧策略
                INetFwPolicy.Rules.Remove(item.Name);

                //添加新策略
                INetFwRule INetFwRule = (INetFwRule)Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FwRule"));
                INetFwRule.Name = item.Name;
                INetFwRule.Description = item.Description;
                INetFwRule.Enabled = item.Enabled;
                INetFwRule.Action = item.Action;
                INetFwRule.Direction = item.Direction;
                INetFwRule.LocalAddresses = item.LocalAddresses;
                INetFwRule.Protocol = item.Protocol;

                string localPorts = string.Join(",", item.LocalPorts.Select(x => x.PortNumber).OrderBy(x=>x));
                INetFwRule.LocalPorts = localPorts;
                INetFwRule.InterfaceTypes = item.InterfaceTypes;

                string remoteAddresses = string.Join(",", item.RemoteAddresses.Select(x => x.IpAddress).OrderBy(x => x));
                INetFwRule.RemoteAddresses = remoteAddresses.Replace(" ","").Replace(",,",",").Replace(",,", ",");
                INetFwRule.RemotePorts = item.RemotePorts;
                INetFwPolicy.Rules.Add(INetFwRule);

                Console.WriteLine("添加规则：" + item.Name + "成功！");
            }

            //打开防火墙管理界面
            Process.Start("mmc.exe", "wf.msc");
        }

        /// <summary>
        /// 检查配置文件，显示帮助信息
        /// </summary>
        /// <returns></returns>
        private bool ConfigFileCheck()
        {
            bool fileExist = File.Exists(fireWallConfigFile);
            if (fileExist == false)
            {
                StringBuilder configSample = new StringBuilder();
                configSample.AppendLine("#################################################################");
                configSample.AppendLine("###############Windows Server Firewall Manager###################");
                configSample.AppendLine("#################################################################");
                configSample.AppendLine("配置文件名：\t" + fireWallConfigFile);
                configSample.AppendLine("文件内容示例:");

                List<NetFwRule> netFwRules = new List<NetFwRule>();
                NetFwRule netFwRule = new NetFwRule();
                netFwRule.Name = "测试规则";
                netFwRule.Description = "这是一个测试规则";

                //默认
                //netFwRule.Enabled = true;
                //netFwRule.Action = NetFwTypeLib.NET_FW_ACTION_.NET_FW_ACTION_ALLOW;
                //netFwRule.Direction = NetFwTypeLib.NET_FW_RULE_DIRECTION_.NET_FW_RULE_DIR_IN;
                //netFwRule.LocalAddresses = "*";
                //netFwRule.Protocol = 6;

                List<Ports> localPorts = new List<Ports>();
                localPorts.Add(new Ports { PortNumber = 1433, Description = "SQL Server" });
                netFwRule.LocalPorts = localPorts;

                //默认
                //netFwRule.InterfaceTypes = "ALL";

                List<RemoteAddress> remoteAddresses = new List<RemoteAddress>();
                remoteAddresses.Add(new RemoteAddress() { IpAddress = "192.168.1.1", Description = "允许的IP地址1" });
                remoteAddresses.Add(new RemoteAddress() { IpAddress = "192.168.1.2", Description = "允许的IP地址2" });
                remoteAddresses.Add(new RemoteAddress() { IpAddress = "192.168.1.3", Description = "允许的IP地址3" });
                remoteAddresses.Add(new RemoteAddress() { IpAddress = "192.168.1.4", Description = "允许的IP地址4" });
                netFwRule.RemoteAddresses = remoteAddresses;

                //默认
                //netFwRule.RemotePorts = "*";

                netFwRules.Add(netFwRule);

                string jsonSample = new JavaScriptSerializer().Serialize(netFwRules);
                configSample.Append(jsonSample);
                Console.WriteLine(configSample.ToString());

                File.WriteAllText(fireWallConfigFile, jsonSample);
                Process.Start("notepad.exe", fireWallConfigFile);

                Console.WriteLine("按回车退出程序！");
                Console.ReadLine();
            }

            return fileExist;
        }
    }
}
