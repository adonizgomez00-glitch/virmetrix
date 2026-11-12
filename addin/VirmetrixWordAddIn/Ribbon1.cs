using System;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;

namespace VirmetrixWordAddIn
{
    public partial class Ribbon1 : OfficeRibbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e) { }

        private void Record(Word.Document doc, string action, string value = "true")
        {
            try
            {
                var props = doc.CustomDocumentProperties;
                try { props["Virmetrix_" + action].Delete(); } catch { }
                props.Add("Virmetrix_" + action, false, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, value);
                doc.Save();
            }
            catch { }
        }

        public void OnStyleSet(OfficeRibbonControl control) { Record(control.Context as Word.Document ?? Globals.ThisAddIn.Application.ActiveDocument, "style_set_applied"); }
        public void OnShare(OfficeRibbonControl control) { Record(control.Context as Word.Document ?? Globals.ThisAddIn.Application.ActiveDocument, "share_electronically"); }
        public void OnCompat(OfficeRibbonControl control) { Record(control.Context as Word.Document ?? Globals.ThisAddIn.Application.ActiveDocument, "compatibility_issues"); }
        public void OnNavigate(OfficeRibbonControl control) { Record(control.Context as Word.Document ?? Globals.ThisAddIn.Application.ActiveDocument, "navigate_location"); }
        public void OnRemoveBg(OfficeRibbonControl control) { Record(control.Context as Word.Document ?? Globals.ThisAddIn.Application.ActiveDocument, "image_background_removed"); }
        public void OnMarks(OfficeRibbonControl control) { Record(control.Context as Word.Document ?? Globals.ThisAddIn.Application.ActiveDocument, "formatting_marks_visible"); }
        public void OnAltFormat(OfficeRibbonControl control) { Record(control.Context as Word.Document ?? Globals.ThisAddIn.Application.ActiveDocument, "alternative_format", "pdf"); }
        public void OnPrint(OfficeRibbonControl control) { Record(control.Context as Word.Document ?? Globals.ThisAddIn.Application.ActiveDocument, "print_settings", "2up"); }
        public void OnFormatPainter(OfficeRibbonControl control) { Record(control.Context as Word.Document ?? Globals.ThisAddIn.Application.ActiveDocument, "format_painter_used"); }
    }
}
EOF
echo done