using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Visio;

namespace Demo.Struct
{
    public class Node : INode
    {
        public Node(Shape _element = null, int _initLevel = 1)
        {
            CurrentLevel = _initLevel;
            Childs = new List<INode>();
            Element = _element;
        }

        public Shape Element { get; set; }
        public ICollection<INode> Childs { get; set; }

        public bool IsLeaf => Childs?.Count > 0;

        public bool IsRoot => Childs?.Count == 0;

        public int CurrentLevel { get; }

        public int Heigth()
        {
            if (IsLeaf)
                return 1;
            else
            {
                int maxHeigth = 0;
                foreach (var child in Childs)
                {
                    var childHeigth = child.Heigth();
                    if (childHeigth > maxHeigth)
                        maxHeigth = childHeigth;
                }

                return maxHeigth + 1;
            }
        }

        public int Weigth()
        {
            if (IsLeaf)
                return 1;
            else
            {
                int maxWeigth = 0;
                foreach (var child in Childs)
                {
                    var childHeigth = child.Weigth();
                    if (childHeigth > maxWeigth)
                        maxWeigth = childHeigth;
                }

                return maxWeigth + 1;
            }
        }

        public void AddNode<T>(T node) where T : INode
        {
            if (Childs != null) {
                Childs.Add(node);
            }
        }

        public INode SearchNode<T>(int id) where T : INode
        {
            if (Element.ID.Equals(id))
                return this;
            else
            {
                foreach (var node in Childs)
                {
                    var foundNode = node.SearchNode<INode>(id);
                    if (foundNode != null)
                        return foundNode;
                }

                return null;
            }
        }

      
     
    }
}
