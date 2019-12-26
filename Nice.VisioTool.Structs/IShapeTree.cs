using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Nice.VisioTool.Structs
{
    interface IShapeTree : ITree
    {
        INode Root { get; set; }   
    }
}
