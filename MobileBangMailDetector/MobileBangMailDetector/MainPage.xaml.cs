using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xamarin.Forms;

namespace MobileBangMailDetector
{
    public partial class MainPage : ContentPage
    {
        public MainPage()
        {
            InitializeComponent();
        }

        protected async override void OnAppearing()
        {
            base.OnAppearing();

            await (Xamarin.Forms.Application.Current as App).SignIn();

            try
            {
                //string filter = "importance eq 'high'";
                //string filter = "importance eq 'high' " +
                //                "and sender/emailaddress/address eq 'rgreen2005@msn.com'";
                string filter = "importance eq 'high' & isread eq 'false' " +
                                "& from/emailaddress/address eq 'rgreen2005@msn.com'";
                IMailFolderMessagesCollectionPage folderMessages =
                    await App.GraphClient.Me.MailFolders.Inbox.Messages.Request()
                                                        .Filter(filter).GetAsync();

                bool messagesFound = false;

                if (folderMessages.Count() > 0)
                {
                    foreach (var message in folderMessages)
                    {
                        if (message.IsRead == false)
                        {
                            messagesFound = true;
                        }
                    }
                }

                if (messagesFound == true)
                {
                    VacationStatusLabel.Text = "Vacation is on hold";
                    VacationStatusLabel.TextColor = Color.Red;
                }
                else
                {
                    VacationStatusLabel.Text = "Vacation is a go";
                    VacationStatusLabel.TextColor = Color.Green;
                }
            }
            catch (ServiceException ex)
            {
                VacationStatusLabel.Text = "Vacation is a maybe. Check Outlook";
            }
        }


    }
}
