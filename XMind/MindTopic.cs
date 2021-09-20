using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XMind2Xls.XMind
{
    class MindTopic
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

        public string structureClass
        {
            get;
            set;
        }

        public MindAttached children
        {
            get;
            set;
        }

       
    }
}
