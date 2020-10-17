using FarmacyControl.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace FarmacyControl
{
    public interface IRepository
    {
        public List<Mrc> GetMrc();
        public void WriteToDb(Mrc mrcs);
    }
}
