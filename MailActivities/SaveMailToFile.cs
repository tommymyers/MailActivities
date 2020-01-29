using System.Activities;
using System.ComponentModel;
using System.Net.Mail;
using System.IO;
using Microsoft.Office.Interop.Outlook;

namespace MailActivities {
    [Description("Saves an email to a file.")]
    public class SaveMailToFile : CodeActivity {
        [Category("Input")]
        [Description("Overwrite if file exists.")]
        public bool Overwrite { get; set; }

        [RequiredArgument]
        [Category("Input")]
        [Description("Email to save.")]
        public InArgument<MailMessage> MailMessage { get; set; }

        [RequiredArgument]
        [Category("Input")]
        [DisplayName("File Path")]
        [Description("Where to save the email. If a directory is given, the email is saved within that directory. Else, it's saved to the given path.")]
        public InArgument<string> FilePath { get; set; }

        private string makeStringFilePathSafe(string str) {
            foreach (char c in Path.GetInvalidFileNameChars()) {
                str = str.Replace(c, '\0');
            }
            return str;
        }

        protected override void Execute(CodeActivityContext context) {
            Application outlook = new Application();
            NameSpace mapi = outlook.GetNamespace("MAPI");
            MailItem mail = mapi.GetItemFromID(MailMessage.Get(context).Headers.Get("UID"));
            string filePathString = FilePath.Get(context);
            if (Directory.Exists(filePathString)) {
                string fileName = mail.ReceivedTime.ToString("ddMMyyyy HHmmss") + " " + makeStringFilePathSafe(mail.Subject) + ".msg";
                filePathString = Path.Combine(filePathString, fileName);
            } else {
                if (!Path.GetFileName(filePathString).EndsWith(".msg")) {
                    filePathString += ".msg";
                }
            }
            if (File.Exists(filePathString) && Overwrite) {
                File.Delete(filePathString);
            }
            if (!File.Exists(filePathString)) {
                mail.SaveAs(filePathString);
            }

            ReleaseComObject(mail);
            ReleaseComObject(mapi);
            ReleaseComObject(outlook);
        }

        private static void ReleaseComObject(object obj) {
            if (obj != null) {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
        }
    }
}
