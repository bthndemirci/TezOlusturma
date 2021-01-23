using System.Windows.Forms;
using DevExpress.XtraSplashScreen;

namespace TezOlusturma.App
{
    public static class DxWaitForm
    {
        public static void Show(Form patternForm)
        {
            SplashScreenManager.ShowForm(patternForm, typeof(FrmWait), true, true, false);
        }

        public static void Close()
        {
            SplashScreenManager.CloseForm(false);
        }
    }
}