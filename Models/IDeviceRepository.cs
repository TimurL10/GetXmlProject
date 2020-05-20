using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using GetXml.Models;

namespace GetXml.Models
{
    public interface IDeviceRepository { 
        void Delete(double id);
        Device Get(double id);
        List<Device> GetDevices();
        void Update(Device device);
        void Add(Device device);
        void AddAddress(Device device);
        List<Device> GetDataForReport();
    }
}
