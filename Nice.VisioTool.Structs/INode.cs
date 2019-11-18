using Microsoft.Office.Interop.Visio;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Demo.Struct
{
    public interface INode : ITree
    {
        Shape Element { get; set; }
        ICollection<INode> Childs { get; set; }

        bool IsLeaf { get; }
        bool IsRoot { get; }
        int CurrentLevel { get; }
        int Heigth();
        int Weigth();
    }
}
