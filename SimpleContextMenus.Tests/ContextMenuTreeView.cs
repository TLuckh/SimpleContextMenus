using System.Text;

namespace SimpleContextMenus.Tests;

// public class ContextMenuTreeView<T>
// {
//     public readonly Node<T> rootNode;
//
//     /// <summary>
//     /// Makes a new ContextMenuTreeView from scratch with the given root value.
//     /// </summary>
//     /// <param name="rootValue"></param>
//     public ContextMenuTreeView(T rootValue)
//     {
//         rootNode = new Node<T>(rootValue);
//     }
//
//     /// <summary>
//     /// Makes a new ContextMenuTreeView with the given root node. Shallow copy!
//     /// </summary>
//     /// <param name="rootNode"></param>
//     public ContextMenuTreeView(Node<T> rootNode)
//     {
//         this.rootNode = rootNode;
//     }
//
//     public override string ToString()
//     {
//         return ToStringInternal(rootNode).ToString();
//
//         StringBuilder ToStringInternal(Node<T> node, string indent = "  ")
//         {
//             var returnString = new StringBuilder();
//             returnString.Append(indent + node);
//             foreach (Node<T> child in node.Children)
//             {
//                 returnString.Append(ToStringInternal(child, indent + indent));
//             }
//
//             return returnString;
//         }
//     }
// }


/// <summary>
/// Helper class to create a tree view of the context menu being built for the tests, and for pretty printing it.
/// </summary>
/// <typeparam name="T"></typeparam>
public class Node<T>
{
    public T Value { get; set; }

    public List<Node<T>> Children { get; } = new List<Node<T>>();

    public Node(T value)
    {
        Value = value;
    }

    public Node<T> AddChild(T childValue)
    {
        var child = new Node<T>(childValue);
        Children.Add(child);
        return child;
    }

    public Node<T> AddChild(Node<T> child)
    {
        Children.Add(child);
        return child;
    }
    
    public override string ToString()
    {
        return ToStringInternal(this).ToString();

        StringBuilder ToStringInternal(Node<T> node, string indent = "  " , string parentItemID = "", int currentItemID = 0)
        {
            
            string idInterior;
            if (parentItemID == "") // base case of recursion
            {
                idInterior = $"{currentItemID}";
            }
            else
            {
                idInterior = $"{parentItemID}.{currentItemID}";
            }
            string id = $"[{idInterior}]".PadRight(11);
                
            var returnString = new StringBuilder();
            returnString.Append(id + indent  + node.Value + "\n");
            
            int childCounter = 0;
            foreach (Node<T> child in node.Children)
            {
                returnString.Append(ToStringInternal(child, indent + indent,$"{idInterior}", childCounter));
                childCounter++;
            }

            return returnString;
        }
    }
}
