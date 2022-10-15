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
                sernder: "alexw@contoso.com", 
                toRecipients: new string[] { "lidiah@contoso.com", "henriettam@contoso.com" },
                subject: "Send mail",
                boodyContent: "Send the <b>message</b> specified in the request body using either JSON or MIME format.", bodyContentType: BodyType.Html,
                attachments: new string[] { @"C:\temp\user_sendMail.pdf" }
                );
        }
    }
}
