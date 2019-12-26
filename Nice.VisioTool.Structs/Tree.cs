using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Nice.VisioTool.Structs
{
    public class Tree : ITree
    {
        public void AddNode<T>(T node) where T : INode
        {

            throw new NotImplementedException();
        }

        public INode SearchNode<T>(int id) where T : INode
        {

            throw new NotImplementedException();
        }
    }
}
