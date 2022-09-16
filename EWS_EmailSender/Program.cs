using System;

namespace EWS_EmailSender
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                //SendEmailByEwsApi.SendEmail();
                SendEmailByEA.SendEmail();
            }
            catch (Exception exc)
            {
                Common.ShowMessage("Error: " + exc.Message, true);
            }
            finally
            {
                Console.ReadLine();
            }
        }
    }
}