using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace GetXml.Models
{
    [Serializable]
    [XmlRoot(ElementName = "device")]
    public class Device
    {
        [XmlAttribute(AttributeName = "id")]
        public double Id { get; set; }

        [XmlAttribute(AttributeName = "name")]
        public string Name { get; set; }

        //[XmlAttribute(AttributeName = "version")]
        //public string Version { get; set; }

        [XmlAttribute(AttributeName = "status")]
        public string Status { get; set; }

        //[XmlAttribute(AttributeName = "os")]
        //public string Os { get; set; }

        [XmlAttribute(AttributeName = "campaign_name")]
        public string Campaign_Name { get; set; }

        [XmlAttribute(AttributeName = "ip")]
        public string Ip { get; set; }

        //[XmlAttribute(AttributeName = "created_at")]
        //public DateTime CreatedAt { get; set; }

        //[XmlAttribute(AttributeName = "connected_at")]
        //public DateTime ConnectedAt { get; set; }

        [XmlAttribute(AttributeName = "last_online")]
        public DateTime Last_Online { get; set; }

        [XmlAttribute(AttributeName = "time_offline")]
        public double Hours_Offline { get; set; }
        public string Address { get; set; }


        public Device()
        {

        }
        public Device(double Id, string Name, string Status, string campaign_name, string Ip, DateTime last_online, double hours_offline, string address)
        {
            this.Id = Id;
            this.Name = Name;
            this.Status = Status;
            this.Campaign_Name = campaign_name;
            this.Ip = Ip;            
            this.Last_Online = last_online.ToLocalTime();
            this.Hours_Offline= hours_offline;
            this.Address = address;
        }

        public Device(string name, string address)
        {
            this.Name = name;
            this.Address = address;
        }
    }

    [XmlRoot(ElementName = "xml")]
    public class Xml
    {
        [XmlElement(ElementName = "device")]
        public List<Device> Devices { get; set; }
    }    
    
}
