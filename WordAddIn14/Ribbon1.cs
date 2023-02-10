using System;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Word;

namespace WordAddIn14
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Header1_Click(object sender, RibbonControlEventArgs e)
        {
            Application wordApp = Globals.ThisAddIn.Application;
            Template buildingBlock = wordApp.Templates["Header1"];
            buildingBlock.BuildingBlockEntries.Item(1).Insert(wordApp.Selection.Range, Type.Missing);
        }
    }
}
