using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Demo.Struct
{
    public interface ITree
    {
        void AddNode<T>(T node) where T : INode;
        INode SearchNode<T>(int id) where T : INode;
    }
}
