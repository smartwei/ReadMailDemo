# ReadMailDemo

#用c#获取outlook(outlook2013/outlook2016)中的邮件  

#关键代码如下  

            Microsoft.Office.Interop.Outlook.Application myOutlookApp = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.NameSpace myNameSpace = myOutlookApp.GetNamespace("MAPI");
 
            Microsoft.Office.Interop.Outlook.MAPIFolder myFolderInbox = myNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);//获取收件箱对象，如获取其他箱可在参数中控制


            Microsoft.Office.Interop.Outlook.Items MailItems = myFolderInbox.Items as Microsoft.Office.Interop.Outlook.Items;
            Console.WriteLine("mail Count:" + MailItems.Count.ToString());
            int showCount = 0;
            for (int index = MailItems.Count; index >0; index--) {
                try
                {
                    //倒序才是从最近的收到的邮件显示
                    Microsoft.Office.Interop.Outlook.MailItem item = MailItems[index] as Microsoft.Office.Interop.Outlook.MailItem;
                    Console.WriteLine("======================================================");
                    Console.WriteLine("Subject:" + item.Subject.ToString());
                    Console.WriteLine("ReceivedTime:" + item.ReceivedTime.ToString());
                    Console.WriteLine("Body:" + item.Body.ToString().Substring(0, 10));

                }
                catch (System.Exception ex)
                {
                    Console.WriteLine("获取失败："+ex.Message);
                }
                showCount++;
                if (showCount > 10) break;//只显示10封
            }
