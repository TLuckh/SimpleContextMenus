using System.Runtime.ExceptionServices;
using SharpShell.Interop;

namespace SimpleContextMenus.Tests;

public class UnitTest1
{
    
    [Fact]
    public void Test1()   // ← kein static, kein [STAThread]
    {
        StaTestHelper.Run(() =>
        {
            var testPaths = new[]
            {
                @"E:\Downloads\Bilder\20260428_132509.jpg",
            };

            foreach (var path in testPaths)
            {
                Console.WriteLine($"\n=== {Path.GetFileName(path)} ===");

                using var tester = new ContextMenuHandlerTester(path);
                var items = tester.ReadAllItems();

                if (items.Count == 0)
                {
                    Console.WriteLine(
                        "  (keine Einträge – konditionaler Handler hat nichts hinzugefügt)");
                    continue;
                }

                PrintItems(items, 0);
            }
        });
    }

    private static void PrintItems(List<ContextMenuItem> items, int indent)
    {
        foreach (var item in items)
        {
            var prefix = new string(' ', indent * 2);
            if (item.IsSeparator)
                Console.WriteLine($"{prefix}---");
            else
            {
                Console.WriteLine(
                    $"{prefix}[{item.CommandId:D3}] " +
                    $"{item.Label,-30} " +
                    $"Verb: {item.Verb}");

                if (item.Children.Count > 0)
                    PrintItems(item.Children, indent + 1);
            }
        }
    }


    
    
    // Hilfsmethode die den Test in einem echten STA-Thread ausführt
    public static class StaTestHelper
    {
        public static void Run(Action action)
        {
            Exception? exception = null;

            var thread = new Thread(() =>
            {
                try   { action(); }
                catch (Exception ex) { exception = ex; }
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            if (exception != null)
                ExceptionDispatchInfo.Capture(exception).Throw();
            // ↑ wirft die Original-Exception mit originalem Stacktrace
        }
    }

}