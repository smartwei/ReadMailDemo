using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace ReadEmailDemo
{
    public partial class Form1 : Form
    {

        List<MyMail> listMail = new List<MyMail>();
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            getOutLookMail3();
        }

        

        private void getOutLookMail3()
        {

            //参考https://www.cnblogs.com/freeliver54/p/10801552.html
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

                    MyMail mail = new MyMail();
                    mail.ID = index;
                    mail.Subject = item.Subject;
                    mail.Body = item.Body;
                    listMail.Add(mail);
                }
                catch (System.Exception ex)
                {
                    Console.WriteLine("获取失败："+ex.Message);
                }
                showCount++;
                if (showCount > 10) break;//只显示10封
            }

            bindListMail();

        }


        private void bindListMail() {
            BindingSource bs = new BindingSource();
            bs.DataSource = listMail;
            listboxMail.DataSource = bs;
            listboxMail.DisplayMember= "Subject";
            listboxMail.ValueMember = "ID";

        }

        private void listboxMail_SelectedIndexChanged(object sender, EventArgs e)
        {
            MyMail mail = (MyMail)listboxMail.SelectedItem;
            txtSubject.Text = mail.Subject;
            txtBody.Text = mail.Body;
            //int nSelectedID = (int)ListBox.SelectedValue;
        }
    }


    public class MyMail {
        public int ID { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
    }
}
