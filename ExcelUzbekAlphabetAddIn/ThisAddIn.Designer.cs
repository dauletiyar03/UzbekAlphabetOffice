namespace ExcelUzbekAlphabetAddIn
{
    [global::Microsoft.VisualStudio.Tools.Applications.Runtime.StartupObjectAttribute(0)]
    public sealed partial class ThisAddIn : global::Microsoft.Office.Tools.AddInBase
    {
        internal global::Microsoft.Office.Interop.Excel.Application Application;
        public ThisAddIn(global::Microsoft.Office.Tools.Excel.ApplicationFactory factory, global::System.IServiceProvider serviceProvider) : base(factory, serviceProvider, "AddIn", "ThisAddIn") { Globals.Factory = factory; }
        protected override void Initialize() { base.Initialize(); Application = GetHostItem<global::Microsoft.Office.Interop.Excel.Application>(typeof(global::Microsoft.Office.Interop.Excel.Application), "Application"); Globals.ThisAddIn = this; }
        protected override void FinishInitialization() { InternalStartup(); OnStartup(); }
    }
}
