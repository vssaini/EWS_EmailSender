using System;
using System.Diagnostics;
using Microsoft.Exchange.WebServices.Data;

namespace EWS_EmailSender
{
    internal class Program
    {
        // This one is 365Dev test account shared by Wunz
        //private const string Username = "AdeleV@4274w4.onmicrosoft.com";
        //private const string Password = "Puw76126";

        private const string Username = "gregtest@kudurrustone.com";
        private const string Password = "Ab5%B2@6";

        private const string RecipientEmail = "vssaini.dev@gmail.com";

        static void Main(string[] args)
        {
            try
            {
                var service = GetExchangeService();
                SendEmail(service);
            }
            catch (Exception exc)
            {
                ShowMessage("Error: " + exc.Message, true);
            }
            finally
            {
                Console.ReadLine();
            }
        }

        private static ExchangeService GetExchangeService()
        {
            Console.WriteLine("Connecting to Exchange Service... It can take 10-40 seconds.");

            var watch = new Stopwatch();
            watch.Start();

            // Instantiating ExchangeService with an empty constructor will create an instance
            // that is bound to the latest known version of Exchange. 
            var service = new ExchangeService
            {
                Credentials = new WebCredentials(Username, Password),

                // For tracing calls
                //TraceEnabled = true,
                //TraceFlags = TraceFlags.All
            };

            service.AutodiscoverUrl(Username, RedirectionUrlValidationCallback);
            watch.Stop();

            if (service.Url != null)
            {
                ShowMessage($"Connected to Exchange Service successfully in {watch.Elapsed.Seconds} seconds");
                Console.WriteLine(Environment.NewLine);
            }
            else
            {
                ShowMessage("Failed to connect to Exchange Service.", true);
            }

            return service;
        }

        private static void SendEmail(ExchangeService service)
        {
            Console.WriteLine($"Sending email to {RecipientEmail}");

            var email = new EmailMessage(service);

            email.ToRecipients.Add(RecipientEmail);

            email.Subject = "EWS";
            email.Body = new MessageBody("This is the first email I've sent by using the EWS Managed API.");

            email.Send();
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            var result = false;
            var redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }

        private static void ShowMessage(string msg, bool error = false)
        {
            Console.ForegroundColor = error ? ConsoleColor.DarkRed : ConsoleColor.Green;
            Console.WriteLine(msg);
            Console.ResetColor();
        }
    }
}