using System;

namespace EWS_EmailSender
{
    public static class Common
    {
        public static string UserEmail { get; set; } = "gregtest@kudurrustone.com";
        public static string Password { get; set; } = "Ab5%B2@6";

        public static string RecipientEmail { get; set; } = "vssaini.dev@gmail.com";

        public static void ShowMessage(string msg, bool error = false)
        {
            Console.ForegroundColor = error ? ConsoleColor.DarkRed : ConsoleColor.Green;
            Console.WriteLine(msg);
            Console.ResetColor();
        }
    }
}
