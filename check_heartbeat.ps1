
Add-Type -TypeDefinition @"
using System;
using System.IO;
using System.Text;

public class MailChecker {
    public static void Run() {
        try {
            dynamic outlook = Activator.CreateInstance(Type.GetTypeFromProgID("Outlook.Application"));
            dynamic ns = outlook.GetNamespace("MAPI");
            dynamic inbox = ns.GetDefaultFolder(6);

            DateTime lastCheck = DateTime.Now.AddMinutes(-30);
            Console.WriteLine("LAST_CHECK: " + lastCheck.ToString("yyyy-MM-dd HH:mm:ss"));

            foreach (dynamic folder in inbox.Folders) {
                if (folder.Name.Contains("\uc911\uc694") || folder.Name.Contains("\uc5b4\ubb34")) {
                    foreach (dynamic item in folder.Items) {
                        try {
                            if (item.ReceivedTime > lastCheck) {
                                Console.WriteLine("NEW_MAIL|FROM:" + (string)item.SenderName + "|SUBJ:" + (string)item.Subject);
                            }
                        } catch {}
                    }
                }
            }
        } catch (Exception e) {
            Console.WriteLine("ERROR: " + e.Message);
        }
    }
}
"@ -ReferencedAssemblies "Microsoft.CSharp"
[MailChecker]::Run()
