using System;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace VirmetrixWordAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, EventArgs e) { }
        private void ThisAddIn_Shutdown(object sender, EventArgs e) { }

        private void InternalStartup()
        {
            this.Startup += ThisAddIn_Startup;
            this.Shutdown += ThisAddIn_Shutdown;
        }
    }
}
EOF
echo done