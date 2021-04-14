using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using GetXml.Models;

namespace GetXml.Models
{
    public interface IDeviceRepository { 
       
        void PostActivity(double id, DateTime date, int active);

    }
}
