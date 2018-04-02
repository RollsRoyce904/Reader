using System;
using System.Collections.Generic;

namespace Reader
{
    class Mapper
    {
        public List<string> Headers { get; set; }
        public int HeaderCount { get; set; }

        public Dictionary<String, IEnumerable<String>> MappedOut { get; set; }

        // make fx to map the MappedOut var

        public Mapper()
        {

        }

    }
}
