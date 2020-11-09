using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
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

        [Display(Name = "Имя")]
        [XmlAttribute(AttributeName = "name")]
        public string Name { get; set; }

        [Display(Name = "Статус")]
        [XmlAttribute(AttributeName = "status")]
        public string Status { get; set; }

        [Display(Name = "Кампейн")]
        [XmlAttribute(AttributeName = "campaign_name")]
        public string Campaign_Name { get; set; }

        [Display(Name = "IP")]
        [XmlAttribute(AttributeName = "ip")]
        public string Ip { get; set; }

        [Display(Name = "Был Online")]
        [XmlAttribute(AttributeName = "last_online")]
        public DateTime Last_Online { get; set; }

        [Display(Name = "Адрес")]
        public string Address { get; set; }       

        [Display(Name = "Заметка")]
        public string Note { get; set; }

        [Display(Name = "Часы Offline")]
        [XmlAttribute(AttributeName = "time_offline")]
        public double Hours_Offline { get; set; }
        
        [Display(Name = "Дни Offline")]
        public double SumHours { get; set; }
        


        public Device()
        {

        }
        public Device(double Id, string Name, string Status, string campaign_name, string Ip, DateTime last_online, string address,  string note, double hours_offline, double sum_hours)
        {
            this.Id = Id;
            this.Name = Name;
            this.Status = Status;
            this.Campaign_Name = campaign_name;
            this.Ip = Ip;
            this.Last_Online = last_online;
            this.Address = address;
            this.Note = note;
            this.Hours_Offline= hours_offline;
            this.SumHours = sum_hours;
        }

        public Device(string name, string address)
        {
            this.Name = name;
            this.Address = address;
        }

        public Device(int id, string note)
        {
            this.Id = id;
            this.Note = note;
        }
    }

    [XmlRoot(ElementName = "xml")]
    public class Xml
    {
        [XmlElement(ElementName = "device")]
        public List<Device> Devices { get; set; }
    }    
    
}
