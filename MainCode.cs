 public partial class ThisAddIn
    {
      
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
          SaveUnreadAttachments();
          this.Application.NewMailEx += app_NewMailEx;
        }

        void app_NewMailEx(string EntryIDCollection)
        {
           // MessageBox.Show("DZFSFSDF");
            try
            {
                SaveUnreadAttachments();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                //проглатываем все ошибки
            }



        }

        private void SaveUnreadAttachments()
        {
            var ns =Application.GetNamespace("MAPI");
            var inbox = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            var unread = inbox.Items.Restrict("[Unread]=TRUE");
            var todayFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            todayFolder = Path.Combine(todayFolder, "Attachments", DateTime.Today.ToShortDateString());
            if (!Directory.Exists(todayFolder))
            {
                Directory.CreateDirectory(todayFolder);
            }
            foreach (object item in unread)
            {
                var mi = item as Outlook.MailItem;
                if (mi != null)
                {
                    foreach (var itm in mi.Attachments)
                    {
                        var att = itm as Outlook.Attachment;
                        if (att != null)
                        {
                            var attfn = Path.Combine(todayFolder, att.FileName);
                            
                            File.Delete(attfn);
                            att.SaveAsFile(attfn);
                        }
                    }
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Код, автоматически созданный VSTO

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
