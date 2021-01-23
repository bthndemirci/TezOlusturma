using System;
using DevExpress.XtraWaitForm;

namespace TezOlusturma.App
{
    public partial class FrmWait : WaitForm
    {
        public FrmWait()
        {
            InitializeComponent();
            progressPanel.AutoHeight = true;
        }

        #region Overrides

        public override void SetCaption(string caption)
        {
            base.SetCaption(caption);
            progressPanel.Caption = caption;
        }
        public override void SetDescription(string description)
        {
            base.SetDescription(description);
            progressPanel.Description = description;
        }
        public override void ProcessCommand(Enum cmd, object arg)
        {
            base.ProcessCommand(cmd, arg);
        }

        #endregion

        public enum WaitFormCommand
        {
        }
    }
}