using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerBIEmbedded_Native.types
{
    public class workSpaceList
    {
        public string Name { get; set; }
        public string Id { get; set; }

        public workSpaceList(string name, string id)
        {
            Name = name;
            Id = id;
        }
    }

    public class PBIContentObject
    {
        public string Name { get; set; }
        public string EmbeddedUrl { get; set; }
        public string EmbeddedId { get; set; }

        public PBIContentObject(string name, string Url, string Id)
        {
            Name = name;
            EmbeddedUrl = Url;
            EmbeddedId = Id;
        }
    }

    public static class PBIObjectType
    {
        public static string Report = "report";
        public static string Dashboard = "dashboard";
        public static string Tile = "tile";
    }

    public enum ViewModeType
    {
        View=0,
        Edit=1,
        Create=2
    }

    public enum TokenType
    {
        Aad=0,
        Embed=1
    }
}
