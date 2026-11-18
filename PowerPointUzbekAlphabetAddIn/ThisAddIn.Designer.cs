using PowerPoint = Microsoft.Office.Interop.PowerPoint;
namespace PowerPointUzbekAlphabetAddIn
{
    [global::Microsoft.VisualStudio.Tools.Applications.Runtime.StartupObjectAttribute(0)]
    public sealed partial class ThisAddIn : global::Microsoft.Office.Tools.AddInBase
    {
        internal PowerPoint.Application Application;
        public ThisAddIn(global::Microsoft.Office.Tools.Factory factory, global::System.IServiceProvider serviceProvider) : base(factory, serviceProvider, "AddIn", "ThisAddIn") { Globals.Factory = factory; }
        protected override void Initialize() { base.Initialize(); Application = GetHostItem<PowerPoint.Application>(typeof(PowerPoint.Application), "Application"); Globals.ThisAddIn = this; }
        protected override void FinishInitialization() { InternalStartup(); OnStartup(); }
    }
}
