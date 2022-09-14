using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestAddIn
{
    class configFile
    {
        public string exportIdentity { get; set; }
        public string registryPath { get; set;  }
        public string dbUser { get; set; }
        public string dbPass { get; set; }
        public string dbName { get; set; }

        public bool debugMode { get; set; }
        
    }
}
