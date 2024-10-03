
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;



namespace WorkWithBlocks
{
    public class ThinBlocks
    {
        [CommandMethod("ThinBlocks")]
        public void ThinChoosenBlocks()     ///убираем блоки накладываемые на другие объекты
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database dB = doc.Database;
            Editor ed = doc.Editor;

            // Starts a new transaction with the Transaction Manager
            using (Transaction trans = dB.TransactionManager.StartTransaction())
            {
             
                trans.Commit();
            }
        }
    }
}