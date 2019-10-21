using OpenQA.Selenium;
using System;
using System.Configuration;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Security.AccessControl;
using System.Security.Principal;

namespace Task2.Helpers
{
    public class CaptureScreenshot
    {
        public void TakeScreenshot(IWebDriver driver, string ScreenshotName, String Folder)
        {
            try
            {
                string saveLocation = System.IO.Path.Combine(ConfigurationManager.AppSettings["ScreenshotLocation"], Folder);
                string TimeStamp = DateTime.Now.ToString(@"yyyy-MM-dd-HH.mm.ss");
                string fullPath = System.IO.Path.Combine(saveLocation, ScreenshotName + "_" + TimeStamp + ".png");
                if (!Directory.Exists(saveLocation))
                    Directory.CreateDirectory(saveLocation);

                DirectoryInfo di = new DirectoryInfo(saveLocation);
                DirectorySecurity dSec = di.GetAccessControl();
                dSec.AddAccessRule(new FileSystemAccessRule(new SecurityIdentifier(WellKnownSidType.WorldSid, null), FileSystemRights.FullControl, InheritanceFlags.ObjectInherit | InheritanceFlags.ContainerInherit, PropagationFlags.NoPropagateInherit, AccessControlType.Allow));
                di.SetAccessControl(dSec);

                ITakesScreenshot sdriver = driver as ITakesScreenshot;
                Screenshot screenshot = sdriver.GetScreenshot();

                using (MemoryStream ms = new MemoryStream(screenshot.AsByteArray))
                using (Image screenShotImage = Image.FromStream(ms))
                {
                    Bitmap cp = new Bitmap(screenShotImage);
                    cp.Save(fullPath, ImageFormat.Png);
                    cp.Dispose();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                throw;
            }
        }
    }
}
