# ReadMailDemo

用c#获取outlook(outlook2013/outlook2016)中的邮件  

关键代码如下  
Microsoft.Office.Interop.Outlook.Application myOutlookApp = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.NameSpace myNameSpace = myOutlookApp.GetNamespace("MAPI");
            //本地邮箱
            Microsoft.Office.Interop.Outlook.MAPIFolder myFolderInbox = myNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);//获取收件箱对象，如获取其他箱可在参数中控制
            //Microsoft.Office.Interop.Outlook.MAPIFolder myFolder = myFolderInbox.Folders["xx"];//“xx”为收件箱下的一个文件夹
            //Microsoft.Office.Interop.Outlook.MAPIFolder MyParentFolder = myFolderInbox.Parent as Microsoft.Office.Interop.Outlook.MAPIFolder;//获取收件箱上一级的文件夹对象，以次来获取与收件箱同级的文件夹
            //Microsoft.Office.Interop.Outlook.MAPIFolder MyNewFolder = MyParentFolder.Folders["yy"];//“yy”为与收件箱同级的文件夹

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
