using System.Runtime.InteropServices;
using System.Text;
using SharpShell.Interop;

namespace SimpleContextMenus.Tests;

// public class ContextMenuTreeView
// {
//     public readonly Node rootNode;
//
//     /// <summary>
//     /// Makes a new ContextMenuTreeView from scratch with the given root value.
//     /// </summary>
//     /// <param name="rootValue"></param>
//     public ContextMenuTreeView(string rootValue)
//     {
//         rootNode = new Node(rootValue);
//     }
//
//     /// <summary>
//     /// Makes a new ContextMenuTreeView with the given root node. Shallow copy!
//     /// </summary>
//     /// <param name="rootNode"></param>
//     public ContextMenuTreeView(Node rootNode)
//     {
//         this.rootNode = rootNode;
//     }
//
//     public override string ToString()
//     {
//         return ToStringInternal(rootNode).ToString();
//
//         StringBuilder ToStringInternal(Node node, string indent = "  ")
//         {
//             var returnString = new StringBuilder();
//             returnString.Append(indent + node);
//             foreach (Node child in node.Children)
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
public class Node
{
    public string ContextMenuEntryName { get; set; }
    public uint ContextMenuEntryHandlerCommandId { get; set; }

    public List<Node> Children { get; } = new List<Node>();

    public Node(string contextMenuEntryName)
    {
        ContextMenuEntryName = contextMenuEntryName;
    }

    public Node AddChild(string childName)
    {
        var child = new Node(childName);
        Children.Add(child);
        return child;
    }

    public Node AddChild(Node child)
    {
        Children.Add(child);
        return child;
    }
    
    /// <summary>
    /// Executes a context menu entry. Needs the contextMenuInterface with which the menu has been built.
    /// </summary>
    /// <param name="contextMenuInterface"></param>
    public void Invoke(IContextMenu contextMenuInterface)
    {
        uint commandId = ContextMenuEntryHandlerCommandId;
        var invoke = new CMINVOKECOMMANDINFO
        {
            cbSize = (uint)Marshal.SizeOf<CMINVOKECOMMANDINFO>(),
            fMask = CMIC.CMIC_MASK_FLAG_NO_UI,
            verb = (IntPtr)(commandId), // 0-basiert
            nShow = 1
        };
    
        IntPtr pici = Marshal.AllocCoTaskMem(Marshal.SizeOf<CMINVOKECOMMANDINFO>());
        try
        {
            Marshal.StructureToPtr(invoke, pici, false);
            contextMenuInterface.InvokeCommand(pici);
        }
        finally
        {
            Marshal.FreeCoTaskMem(pici); // immer freigeben, auch bei Exception
        }
    }



    public override string ToString()
    {
        return ToStringInternal(this).ToString();

        StringBuilder ToStringInternal(Node node, string indent = "  ", string parentItemID = "",
            int currentItemID = 0)
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
            returnString.Append(id + indent + node.ContextMenuEntryName + "\n");

            int childCounter = 0;
            foreach (Node child in node.Children)
            {
                returnString.Append(ToStringInternal(child, indent + indent, $"{idInterior}", childCounter));
                childCounter++;
            }

            return returnString;
        }
    }
}

