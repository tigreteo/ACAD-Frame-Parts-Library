using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.PlottingServices;

namespace FramePartsLibrary
{
    class Plot_Engine
    {
        // A PlotEngine does the actual plotting
        // (can also create one for Preview)
        //***NOTE- always be sure that back ground plotting is off, in code and the users computer.
        public static void plotEngine(PlotInfo pi, string directory, Document doc, Editor ed, string fileNameScheme)
        {
            if (PlotFactory.ProcessPlotState == ProcessPlotState.NotPlotting)
            {
                using (PlotEngine pe = PlotFactory.CreatePublishEngine())
                {
                    // Create a Progress Dialog to provide info or allow the user to cancel
                    PlotProgressDialog ppd = new PlotProgressDialog(false, 1, true);
                    using (ppd)
                    {
                        ppd.set_PlotMsgString(PlotMessageIndex.DialogTitle, "Custom Plot Progress");
                        ppd.set_PlotMsgString(PlotMessageIndex.CancelJobButtonMessage, "Cancel Job");
                        ppd.set_PlotMsgString(PlotMessageIndex.CancelSheetButtonMessage, "Cancel Sheet");
                        ppd.set_PlotMsgString(PlotMessageIndex.SheetSetProgressCaption, "Sheet Set Progress");
                        ppd.set_PlotMsgString(PlotMessageIndex.SheetProgressCaption, "Sheet Progress");
                        ppd.LowerPlotProgressRange = 0;
                        ppd.UpperPlotProgressRange = 100;
                        ppd.PlotProgressPos = 0;

                        // Let's start the plot, at last
                        ppd.OnBeginPlot();
                        ppd.IsVisible = true;
                        pe.BeginPlot(ppd, null);

                        // We'll be plotting a single document
                        pe.BeginDocument(pi, doc.Name, null, 1, true, directory + fileNameScheme);

                        // Which contains a single sheet
                        ppd.OnBeginSheet();
                        ppd.LowerSheetProgressRange = 0;
                        ppd.UpperSheetProgressRange = 100;
                        ppd.SheetProgressPos = 0;

                        PlotPageInfo ppi = new PlotPageInfo();
                        pe.BeginPage(ppi, pi, true, null);
                        pe.BeginGenerateGraphics(null);
                        pe.EndGenerateGraphics(null);

                        // Finish the sheet
                        pe.EndPage(null);
                        ppd.SheetProgressPos = 100;
                        ppd.OnEndSheet();

                        // Finish the document
                        pe.EndDocument(null);

                        // And finish the plot
                        ppd.PlotProgressPos = 100;
                        ppd.OnEndPlot();
                        pe.EndPlot(null);
                    }
                }
            }

            else
            {
                ed.WriteMessage("\nAnother plot is in progress.");
            }
        }
    }
}
