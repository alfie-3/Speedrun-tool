using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NicoSpeedrunTool {
    internal static class Tools {
        public static void WriteMessage(string message, ConsoleColor textColour = ConsoleColor.White) {
            Console.ForegroundColor = textColour;
            Console.WriteLine(message);
            Console.ResetColor();
        }
    }
}
