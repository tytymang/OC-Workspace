
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
                // '중요 업무' (EUC-KR 깨짐 대응: '߿ ' 또는 포함 문자열 확인)
                // 직접 인덱스로 접근하거나 이름을 비교
                if (folder.Name.Contains("\uc911\uc694") || folder.Name.Contains("\uc5b4\ubb34") || folder.Name.Length == 5) {
                    targetFolder = folder;
                    // 상세 확인을 위해 이름 출력
                    Console.WriteLine("CHECKING_FOLDER: " + (string)folder.Name);
                    
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
                }
            }
        } catch (Exception e) {
            Console.WriteLine("ERROR: " + e.Message);
        }
    }
}
"@ -ReferencedAssemblies "Microsoft.CSharp"
[MailGetter]::Run()
