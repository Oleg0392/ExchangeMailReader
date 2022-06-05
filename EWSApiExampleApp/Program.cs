using System;
using System.IO;
using Microsoft.Exchange.WebServices.Data;
using System.Timers;

namespace EWSApiExampleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Initializing application configuration...");
            AppConfig appConfig = new AppConfig("appConfig.xml");

            Console.WriteLine("Initializing Exchange Web Service...");
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);

            if (appConfig.webCredentialsType == 1)
            {
                service.Credentials = new WebCredentials();   //default credentials (if programm in user machine)
            }
            else service.Credentials = new WebCredentials(appConfig.username, appConfig.password, appConfig.domain);
            

            service.Url = new Uri(appConfig.serviceUrl);
            Console.WriteLine("Redirection Url Validation...");
            Console.WriteLine($"Validation: {RedirectionUrlValidationCallback(service.Url.OriginalString)}");

            /*service.TraceEnabled = true;
            service.TraceFlags = TraceFlags.All;*/

            EmailMessage emailMessage = null;
            Console.WriteLine("Creating and preparing ItemView element...");
            ItemView view = new ItemView(appConfig.viewPageSize);
            view.PropertySet = new PropertySet(ItemSchema.Subject, ItemSchema.DateTimeReceived, EmailMessageSchema.IsRead);
            view.Traversal = ItemTraversal.Shallow;
            view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending);

            try
            {
                Console.WriteLine("Try to find elements in mailbox...");
                FindItemsResults<Item> results = service.FindItems(new FolderId(WellKnownFolderName.Inbox, appConfig.userEmailAddress), view);
                Console.WriteLine("Loading properties of finded elements...");
                service.LoadPropertiesForItems(results.Items, PropertySet.FirstClassProperties);
                emailMessage = results.Items[0] as EmailMessage;
                
                /*foreach (Item item in results)
                {
                    Console.WriteLine($"Message subject: {item.Subject}");
                    //Console.WriteLine($"Message id is: {item.Id.ToString()}");
                    if (item is EmailMessage)
                    {
                        emailMessage = item as EmailMessage;
                        Console.WriteLine($"Is Read: {message.IsRead.ToString()}");
                    }
                }*/
            }
            catch (Exception e)
            {
                Console.WriteLine($"{appConfig.LogExceptionMessage1} {e.Message}");
            }



            if (emailMessage is not null)
            {
                
                string filePath = $"{appConfig.parrentFolderPath}{emailMessage.Subject}{appConfig.fileExtension}";
                Console.WriteLine("Saving message body to: {0}",filePath);
                FileStream fileStream = File.Create(filePath);
                fileStream.Close();
                StreamWriter streamWriter = new StreamWriter(filePath);
                streamWriter.WriteLine(emailMessage.Body.Text);
                streamWriter.Close();
                Console.WriteLine("Message body is writted in file.");
            }

            /*EmailMessage emailMessage = new EmailMessage(service);
            emailMessage.ToRecipients.Add("rparobot140@mmk.ru");
            emailMessage.Subject = "testMessage";
            emailMessage.Body = new MessageBody("Hello pridurchatyi durachek, etoya, EWS API)))");
            emailMessage.Send();*/

            Console.ReadLine();
            

        }


        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }

            return result;
        }
    }


}
