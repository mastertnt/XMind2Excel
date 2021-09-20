using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XMind2Xls.XMind
{
    class MindSheet
    {
        public string id
        {
            get;
            set;
        }

        public string title
        {
            get;
            set;
        }

        public MindTopic rootTopic
        {
            get;
            set;
        }
    }
}
