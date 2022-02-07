using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenBots.Studio.Utilities.Documentation
{
    public class RelationDictionary
    {
        public static Dictionary<string, int> CommandTree = new Dictionary<string, int>()
        {
            { "data" , 1 },
            { "loop" , 1 },
            { "if" , 1 }
        };
    }
}
