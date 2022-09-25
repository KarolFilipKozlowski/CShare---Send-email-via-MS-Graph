using Microsoft.Graph;
using MSGraphMailer.GraphHelper;

namespace MSGraphMailer
{
    internal class Program
    {
        static void Main()
        {
            var graphClient = new GetGraphServiceClient(AuthorizedBy.byClientSecret);
            UserSendMail.SendMail(graphClient.GraphClient, 
                sernder: "alexw@4ql5ky.contoso.com", 
                toRecipients: new string[] { "lidiah@4ql5ky.contoso.com", "henriettam@4ql5ky.contoso.com" },
                subject: "Send mail",
                boodyContent: "Send the <b>message</b> specified in the request body using either JSON or MIME format.", bodyContentType: BodyType.Html,
                attachments: new string[] { @"C:\temp\user_sendMail.pdf" }
                );
        }
    }
}