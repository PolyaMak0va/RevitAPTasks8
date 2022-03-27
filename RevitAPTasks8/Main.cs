using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPTasks8
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand // IExternalCommand явл-ся интерфейсом
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            using (var ts = new Transaction(doc, "Export IFC"))
            {
                ts.Start();

                ViewPlan viewPlan = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewPlan))
                    .Cast<ViewPlan>()
                    .FirstOrDefault(v => v.ViewType == ViewType.FloorPlan &&
                                         v.Name.Equals("Level 1"));

                var ifcOption = new IFCExportOptions();

                // укзываем место для сохранения (здесь - рабочий стол)
                doc.Export(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "export.ifc", ifcOption);

                ts.Commit();
            }
            TaskDialog.Show("Сообщение", $"Экспорт в формат IFC выполнен.{ Environment.NewLine}Файл на сохранён на рабочем столе");
            return Result.Succeeded;
        }
        public void BatchPrint(Document doc)
        {
            var sheets = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .ToList();

            var groupedSheets = sheets.GroupBy(sheet => doc.GetElement(new FilteredElementCollector(doc, sheet.Id)
                                                           .OfCategory(BuiltInCategory.OST_TitleBlocks)
                                                           .FirstElementId()).Name);

            var viewSets = new List<ViewSet>();

            PrintManager printManager = doc.PrintManager;

            printManager.SelectNewPrintDriver("PDF-XChange 5.0 for ABBYY FineReader 14");
            printManager.PrintRange = PrintRange.Select;

            ViewSheetSetting viewSheetSetting = printManager.ViewSheetSetting;

            foreach (var groupedSheet in groupedSheets)
            {
                if (groupedSheet.Key == null)
                    continue;

                var viewSet = new ViewSet();

                var sheetsOfGroup = groupedSheet.Select(s => s).ToList();

                foreach (var sheet in sheetsOfGroup)
                {
                    viewSet.Insert(sheet);
                }

                viewSets.Add(viewSet);

                viewSheetSetting.CurrentViewSheetSet.Views = viewSet;

                using (var ts = new Transaction(doc, "Create view set"))
                {
                    ts.Start();
                    viewSheetSetting.SaveAs($"{groupedSheet.Key}_{Guid.NewGuid()}");
                    ts.Commit();
                }

                bool isFormatSelected = false;

                foreach (PaperSize paperSize in printManager.PaperSizes)
                {

                    if (string.Equals(groupedSheet.Key, "А4К") &&
                        string.Equals(paperSize.Name, "A4"))
                    {
                        printManager.PrintSetup.CurrentPrintSetting.PrintParameters.PaperSize = paperSize;
                        printManager.PrintSetup.CurrentPrintSetting.PrintParameters.PageOrientation = PageOrientationType.Portrait;
                        isFormatSelected = true;
                    }

                    else if (string.Equals(groupedSheet.Key, "А3А") &&
                        string.Equals(paperSize.Name, "A3"))
                    {
                        printManager.PrintSetup.CurrentPrintSetting.PrintParameters.PaperSize = paperSize;
                        printManager.PrintSetup.CurrentPrintSetting.PrintParameters.PageOrientation = PageOrientationType.Landscape;
                        isFormatSelected = true;
                    }
                }

                if (!isFormatSelected)
                {
                    TaskDialog.Show("Ошибка!", "Не найден формат");
                    return;
                }
                printManager.CombinedFile = true;
                printManager.SubmitPrint();
            }
        }
    }
}
