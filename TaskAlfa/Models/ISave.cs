using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TaskAlfa.Models
{
    public interface ISave
    {
        void Save(string path, List<string> result);
    }
}
