
Add-Type -TypeDefinition @"
using System;
using System.IO;
using System.Text;

public class MailGetter {
    public static void Run() {
        try {
            dynamic outlook = Activator.CreateInstance(Type.GetTypeFromProgID("Outlook.Application"));
            dynamic ns = outlook.GetNamespace("MAPI");
            dynamic inbox = ns.GetDefaultFolder(6);
            dynamic targetFolder = null;

            foreach (dynamic folder in inbox.Folders) {
                if (folder.Name.Contains("\uc911\uc694 \uc5b4\ubb34")) {
                    targetFolder = folder;
                    break;
                }
            }

            if (targetFolder == null) {
                Console.WriteLine("FOLDER_NOT_FOUND");
                return;
            }

            string savePath = @"C:\Users\307984\.openclaw\workspace\temp_attachments";
            if (!Directory.Exists(savePath)) Directory.CreateDirectory(savePath);

            foreach (dynamic item in targetFolder.Items) {
                try {
                    string sender = item.SenderName;
                    if (sender.Contains("\uae40\ud558\uc601") || sender.Contains("\uc774\uc218\uc815")) {
                        Console.WriteLine("FOUND: " + (string)item.Subject);
                        foreach (dynamic at in item.Attachments) {
                            string fullPath = Path.Combine(savePath, (string)at.FileName);
                            at.SaveAsFile(fullPath);
                            Console.WriteLine("SAVED: " + fullPath);
                        }
                    }
                } catch {}
            }
        } catch (Exception e) {
            Console.WriteLine("ERROR: " + e.Message);
        }
    }
}
"@ -ReferencedAssemblies "Microsoft.CSharp"
[MailGetter]::Run()
