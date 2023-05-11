using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Drawing.Printing;
using System.Threading.Tasks;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace FramePartsLibrary
{
    public struct FrameRecord
    {
        public FrameRecord(string StyleID, DateTime lastModified, string fullPath)
        {
            styleID = StyleID;
            date = lastModified;
            path = fullPath;

            //derive name parts
            //find expected styleID and group number from filename
            StringBuilder groupName = new StringBuilder();
            StringBuilder styleName = new StringBuilder();

            #region find name
            string fileName = styleID;            
            bool secondPart = false;
            bool firstPart = false;
            char[] groupList = { ' ', '-', 'A', 'C' };
            char[] styleList = { ' ', 'S', 'N' };//might need to change this to be more general
                                                 //it assumes and S as in SAM1, or N as in NESTERWOOD

            //loop through file name to generate parts of name
            //first to last
            foreach (char c in fileName)
            {
                //if number is false it isnt complete
                if (!secondPart)
                {
                    if (!firstPart)
                    {
                        //if part isnt one of the separaters add it to end of part
                        if (Array.Exists(groupList, element => element == c))
                        {
                            //if it is a separater then the first part is complete(true)
                            firstPart = true;
                        }
                        else
                            groupName.Append(c);
                    }

                    else {
                        //if par isnt one of the expected closers then add it to the end of the part
                        if (Array.Exists(styleList, element => element == c))
                        {
                            //if it is an expected closer then set second part to true
                            secondPart = true;
                        }
                        else
                            styleName.Append(c);
                    }
                }
            }
            #endregion

            group = groupName.ToString();
            style = styleName.ToString();
        }

        public string path { get; }
        public string group { get; }
        public string style { get; }
        public string styleID { get; }
        public DateTime date { get; }
    }

    public static class Class1
    {
        //create lists of both DB of data that results in a same naming scheme
        //compare lists finding:
        //(1) missing style data from one location to the other
        //(2) data being out of date of the source data
        //Foreach file location on list
        //empty destination folder if exists, create destination folder if doesnt
        //open file run PDF of each spec, saving to destination folder

        [DllImport("accore.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "acedTrans")]
        static extern int acedTrans(double[] point, IntPtr fromRb, IntPtr toRb, int disp, double[] result);

        [CommandMethod("FrameToCatalog", CommandFlags.Session)]
        static public void ProcessPDFs()
        {
            //PDF Locations
            string PDFpath = @"Y:\CNC Files\Parts Catalog";
            //string PDFpath = @"C:\temp\Parts Catalog";

            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.CurrentDocument;
            Editor ed = doc.Editor;

            // needs to have background plotting turned off to run
            //save current plotting setting to return to when finished
            short bgPlot = (short)Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("BACKGROUNDPLOT");
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("BACKGROUNDPLOT", 0);

            //request if this is updating current view, a selection of files/folders, or all
            #region get fileRange
            //set up prompt message
            PromptKeywordOptions pko = new PromptKeywordOptions("");
            pko.Message = ("What range to process?");
            pko.Keywords.Add("All");
            pko.Keywords.Add("Specific Folder");
            pko.Keywords.Add("Current");
            pko.Keywords.Default = "Specific Folder";
            pko.AllowNone = false;
            pko.AppendKeywordsToMessage = true;

            string searchFolder = "";
            List<string> DWGpaths = new List<string>();
            IEnumerable<FileInfo> fileListDWG = null;

            //request from user 
            PromptResult pkeyRes = ed.GetKeywords(pko);
            if (pkeyRes.Status == PromptStatus.OK)
            {
                if (pkeyRes.StringResult.ToUpper() == "ALL")
                {

                    DWGpaths.Add(@"Y:\Engineering\Drawings\Frames");
                    DWGpaths.Add(@"Y:\Engineering\Drawings\Frames\18mm Frames");
                    DWGpaths.Add(@"Y:\Engineering\Drawings\Frames\Bed Drawings");

                    #region get fileInfo from  DWGS
                    //get file info for the frames
                    foreach (string path in DWGpaths)
                    {
                        DirectoryInfo dirDWG = new DirectoryInfo(path);
                        IEnumerable<DirectoryInfo> folderList = dirDWG.GetDirectories("*.*", SearchOption.TopDirectoryOnly);

                        //loop through each folder inside each main folder of frames
                        foreach (DirectoryInfo folder in folderList)
                        {
                            IEnumerable<FileInfo> fileList = folder.GetFiles("*.*", SearchOption.TopDirectoryOnly);
                            //filter for DWGs
                            fileList =
                                from file in fileList
                                where file.Extension == ".dwg" &&
                                !file.Name.Contains("3D") &&
                                !file.Name.Contains("recover")
                                orderby file.DirectoryName
                                select file;

                            //add to main IEnumeration
                            if (fileListDWG == null)
                                fileListDWG = fileList;
                            else
                                fileListDWG = fileListDWG.Concat(fileList);
                        }
                    }
                    #endregion
                }
                else if (pkeyRes.StringResult.ToUpper() == "SPECIFIC")
                {
                    FolderBrowserDialog folder = new FolderBrowserDialog();
                    folder.SelectedPath = @"Y:\Product Development\Style Specifications\";
                    folder.RootFolder = Environment.SpecialFolder.Desktop;

                    if (folder.ShowDialog() == DialogResult.OK)
                    { searchFolder = folder.SelectedPath; }

                    #region get fileInfo from DWGs

                    DirectoryInfo dir = new DirectoryInfo(searchFolder);
                    IEnumerable<FileInfo> fileList = dir.GetFiles("*.*", SearchOption.AllDirectories);
                    //filter for DWGs
                    fileList =
                        from file in fileList
                        where file.Extension == ".dwg" &&
                        !file.Name.Contains("3D")
                        orderby file.DirectoryName
                        select file;

                    //add to main IEnumeration
                    if (fileListDWG == null)
                        fileListDWG = fileList;
                    else
                        fileListDWG = fileListDWG.Concat(fileList);
                    #endregion
                }
                else if (pkeyRes.StringResult.ToUpper() == "CURRENT")
                {
                    //create record of current doc and pass it along
                    FrameRecord curr = new FrameRecord(
                        Path.GetFileNameWithoutExtension(doc.Name),
                        DateTime.Now,
                        doc.Name);
                    processFile(curr, PDFpath, false);
                    return;   
                }

            }
            else
                return;
            #endregion

            #region get fileInfo from PDFS
            //get file info on PDF folder
            DirectoryInfo dirPDF = new DirectoryInfo(PDFpath);
            IEnumerable<FileInfo> fileListPDF = dirPDF.GetFiles("*.*", SearchOption.AllDirectories);
            //filter for pdfs
            fileListPDF =
                from file in fileListPDF
                where file.Extension == ".pdf" || file.Extension == ".txt"
                orderby file.DirectoryName
                select file;
            #endregion

            #region convert fileinfo into struct
            //convert enumerations into lists with same naming scheme
            List<FrameRecord> listPDF = new List<FrameRecord>();
            foreach (FileInfo fi in fileListPDF)
            {
                //get styleID from parent                
                //string[] nameParts = (fi.Name).Split('^');
                string[] pathParts = (fi.DirectoryName).Split('\\');
                FrameRecord newName = new FrameRecord(
                    pathParts[pathParts.Length - 1].ToString(),
                    fi.LastWriteTime,
                    fi.FullName);

                listPDF.Add(newName);                
                
            }
            List<FrameRecord> listDWG = new List<FrameRecord>() ;
            foreach(FileInfo fi in fileListDWG)
            {
                FrameRecord newName = new FrameRecord(
                    Path.GetFileNameWithoutExtension(fi.Name),
                    fi.LastWriteTime,
                    fi.FullName);
                listDWG.Add(newName);
            }
            #endregion

            //compare lists
            //if drawing exists but PDF doesnt, then add to update/create list
            //if drawing is newer than the PDF, then add to update/create list

            //could do with nested for loops but LINQ should be more effeceint
            ////will always find all files to be exempt from other list, need to compare elements of lists
            //List<FrameRecord> createList = listDWG.Except(listPDF).ToList();          
            var createList = from first in listDWG
                             where !listPDF.Any(x => x.styleID == first.styleID)
                             select first;

            //createlist of shared styleIDs
            var updateList = from dwgs in listDWG
                             where listPDF.Any(x => x.styleID == dwgs.styleID && DateTime.Compare(x.date,dwgs.date) < 0)
                             select dwgs;

            //for all in list:
            //open file
            //find all relevant blocks
            //create savename using block info
            foreach (FrameRecord rec in updateList)
            {
                processFile(rec, PDFpath);
            }
            foreach (FrameRecord rec in createList)
            {
                processFile(rec, PDFpath);
            }

            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("BACKGROUNDPLOT", bgPlot);
        }

        private static void processFile(FrameRecord rec, string PDFpath, bool load = true)
        {
            DocumentCollection docMgr = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
            if (load)
            {
                //open file          
                if (File.Exists(rec.path))
                { docMgr.Open(rec.path, true); }
                else
                { docMgr.MdiActiveDocument.Editor.WriteMessage("File " + rec.path + " does not exist."); }
            }

            //switch focus of hosting app??
            //Application.DocumentManager.DocumentActivationEnabled = true;

            //verify/create path to save loc
            string pdfDirectory = PDFpath + @"\" + rec.group + @"\" + rec.styleID + @"\";
            //if exists delete file contents of folder, leave sub directors for anything in folder "old"
            if(Directory.Exists(pdfDirectory))
            {
                DirectoryInfo di = new DirectoryInfo(pdfDirectory);
                foreach(FileInfo file in di.GetFiles())
                { file.Delete(); }
            }
            else
            { Directory.CreateDirectory(pdfDirectory);}//create folder

            //find all blockrefs(all versions of blks)
            //using (DocumentLock docLock = doc.LockDocument())
            Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.CurrentDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            HostApplicationServices.WorkingDatabase = db;
            bool errorLogFlag = false;
            bool specsFound = false;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {                
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                //List<BlockReference> blkCol = new List<BlockReference>();
                foreach (ObjectId msId in btr)
                {
                    if (msId.ObjectClass.DxfName.ToUpper() == "INSERT")
                    {
                        BlockReference blkRef = tr.GetObject(msId, OpenMode.ForRead) as BlockReference;                        
                        string blockName = "";
                        try
                        { blockName = blkRef.Name; }
                        catch (System.Exception)
                        { }
                        if (blockName.Contains("CNC_BORDER"))
                        {
                            //flag that parts do exist
                            specsFound = true;

                            //get naming info from attribute collection
                            string fileName = getPDFName(blkRef, tr);

                            //get window of blkref
                            Extents2d window = coordinates(blkRef);

                            try
                            {
                                //get plotinfo
                                plotSetUp(
                                    window, tr, db, true, true,
                                    pdfDirectory, doc, ed, fileName);
                            }
                            catch
                            {
                                //error log
                                errorLogFlag = true;
                            }
                        }                        
                    }
                }
                //switch focus of host app?

                tr.Commit();
            }
            
            //print out error log if there was a failure
            if(errorLogFlag)
            {
                File.AppendAllText(pdfDirectory + "Plot Fail.txt", "Error trying to print to file");
                //list of issues further up in folders     
                //replace whole filepath with PDFpath + \File Fail List.txt 
                File.AppendAllText(@"Y:\CNC Files\Parts Catalog\Plot Fail List.txt", "Error trying to print to file " + rec.styleID + Environment.NewLine);                
            }           

            //print out a blank if no specs were found
            if(!specsFound)
            {
                File.AppendAllText(pdfDirectory + "No Parts.txt", "No parts found");
                //list of blanks further up in folders    
                File.AppendAllText(@"Y:\CNC Files\Parts Catalog\No Parts List.txt", "No parts found " + rec.styleID + Environment.NewLine);
            }

            //close file
            if(load)
            { docMgr.MdiActiveDocument.CloseAndDiscard(); }
            
        }

        //returned name is based on info in form, could be different from fileName!!
        private static string getPDFName(BlockReference blkRef, Transaction tr)
        {
            AttributeCollection atts = blkRef.AttributeCollection;
            string[] nameParts = new string[4];

            //iterate through the attributes to find the parts for a name
            foreach (ObjectId ID in atts)
            {
                using (DBObject dbObj = tr.GetObject(ID, OpenMode.ForRead) as DBObject)
                {
                    AttributeReference attRef = dbObj as AttributeReference;

                    if (attRef.Tag.Contains("SUITE"))
                    { nameParts[0] = attRef.TextString.Trim(); }
                    else if (attRef.Tag.Contains("ITEM"))
                    { nameParts[1] = attRef.TextString.Trim(); }
                    else if (attRef.Tag.Contains("PART-NO"))
                    { nameParts[2] = attRef.TextString.Trim(); }
                    else if (attRef.Tag.Contains("PARTNAME"))
                    { nameParts[3] = attRef.TextString.Trim(); }
                }
            }

            //might add the version to the end
            string name = String.Format("{0}-{1}-{2}-{3}",
                nameParts[0], nameParts[1], nameParts[2], nameParts[3]);

            return name;
        }

        //acquire the extents of the frame and convert them from UCS to DCS, in case of view rotation
        static public Extents2d coordinates(BlockReference blkRef)
        {
            Extents3d ext = (Extents3d)blkRef.Bounds;
            Point3d firstInput = ext.MaxPoint;
            Point3d secondInput = ext.MinPoint;

            double minX;
            double minY;
            double maxX;
            double maxY;

            //sort through the values to be sure that the correct first and second are assigned
            if (firstInput.X < secondInput.X)
            { minX = firstInput.X; maxX = secondInput.X; }
            else
            { maxX = firstInput.X; minX = secondInput.X; }

            if (firstInput.Y < secondInput.Y)
            { minY = firstInput.Y; maxY = secondInput.Y; }
            else
            { maxY = firstInput.Y; minY = secondInput.Y; }


            Point3d first = new Point3d(minX, minY, 0);
            Point3d second = new Point3d(maxX, maxY, 0);
            //converting numbers to something the system uses (DCS) instead of UCS
            ResultBuffer rbFrom = new ResultBuffer(new TypedValue(5003, 1)), rbTo = new ResultBuffer(new TypedValue(5003, 2));
            double[] firres = new double[] { 0, 0, 0 };
            double[] secres = new double[] { 0, 0, 0 };
            //convert points
            acedTrans(first.ToArray(), rbFrom.UnmanagedObject, rbTo.UnmanagedObject, 0, firres);
            acedTrans(second.ToArray(), rbFrom.UnmanagedObject, rbTo.UnmanagedObject, 0, secres);
            Extents2d window = new Extents2d(firres[0], firres[1], secres[0], secres[1]);

            return window;
        }

        //determine if the plot is landscape or portrait based on which side is longer
        static public PlotRotation orientation(Extents2d ext)
        {
            PlotRotation portrait = PlotRotation.Degrees180;
            PlotRotation landscape = PlotRotation.Degrees270;
            double width = ext.MinPoint.X - ext.MaxPoint.X;
            double height = ext.MinPoint.Y - ext.MaxPoint.Y;
            if (Math.Abs(width) > Math.Abs(height))
            { return landscape; }
            else
            { return portrait; }
        }

        //set up plotinfo
        static public void plotSetUp(Extents2d window, Transaction tr, Database db, bool scaleToFit, bool pdfout,
            string pdfDirectory, Document doc, Editor ed, string fileName)
        {

            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
            // We need a PlotInfo object linked to the layout
            PlotInfo pi = new PlotInfo();
            pi.Layout = btr.LayoutId;

            // Reference the Layout Manager
            //LayoutManager acLayoutMgr = LayoutManager.Current;

            //current layout
            Layout lo = (Layout)tr.GetObject(btr.LayoutId, OpenMode.ForRead);
            //Layout lo = tr.GetObject(LayoutManager.Current.GetLayoutId(LayoutManager.Current.CurrentLayout), OpenMode.ForRead) as Layout;
            //Layout lo = tr.GetObject(acLayoutMgr.GetLayoutId(acLayoutMgr.CurrentLayout), OpenMode.ForRead) as Layout;

            // We need a PlotSettings object based on the layout settings which we then customize
            PlotSettings ps = new PlotSettings(lo.ModelType);
            //PlotSettings ps = new PlotSettings(false);
            ps.CopyFrom(lo);

            //The PlotSettingsValidator helps create a valid PlotSettings object
            PlotSettingsValidator psv = PlotSettingsValidator.Current;
            psv.RefreshLists(ps);

            //set rotation
            psv.SetPlotRotation(ps, orientation(window));

                
            // We'll plot the window, centered, scaled, landscape rotation
            psv.SetPlotType(ps, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);
            psv.SetPlotWindowArea(ps, window);

            // Set the plot scale
            psv.SetUseStandardScale(ps, true);
            if (scaleToFit == true)
            { psv.SetStdScaleType(ps, StdScaleType.ScaleToFit); }
            else
            { psv.SetStdScaleType(ps, StdScaleType.StdScale1To1); }

            // Center the plot
            psv.SetPlotCentered(ps, true);//finding best location

            //get printerName from system settings
            PrinterSettings settings = new PrinterSettings();
            string defaultPrinterName = settings.PrinterName;

            psv.RefreshLists(ps);
            // Set Plot device & page size 
            // if PDF set it up for some PDF plotter
            if (pdfout == true)
            {
                psv.SetPlotConfigurationName(ps, "DWG to PDF.pc3", null);
                var mns = psv.GetCanonicalMediaNameList(ps);
                if (mns.Contains("ANSI_expand_A_(8.50_x_11.00_Inches)"))
                { psv.SetCanonicalMediaName(ps, "ANSI_expand_A_(8.50_x_11.00_Inches)"); }
                else
                { string mediaName = setClosestMediaName(psv, ps, 8.5, 11, true); }
            }
            else
            {
                psv.SetPlotConfigurationName(ps, defaultPrinterName, null);
                var mns = psv.GetCanonicalMediaNameList(ps);
                if (mns.Contains("Letter"))
                { psv.SetCanonicalMediaName(ps, "Letter"); }
                else
                { string mediaName = setClosestMediaName(psv, ps, 8.5, 11, true); }
            }

            //rebuilts plotter, plot style, and canonical media lists
            //(must be called before setting the plot style)
            psv.RefreshLists(ps);

            //ps.ShadePlot = PlotSettingsShadePlotType.AsDisplayed;
            //ps.ShadePlotResLevel = ShadePlotResLevel.Normal;

            //plot options
            //ps.PrintLineweights = true;
            //ps.PlotTransparency = false;
            //ps.PlotPlotStyles = true;
            //ps.DrawViewportsFirst = true;
            //ps.CurrentStyleSheet

            // Use only on named layouts - Hide paperspace objects option
            // ps.PlotHidden = true;

            //psv.SetPlotRotation(ps, PlotRotation.Degrees180);


            //plot table needs to be the custom heavy lineweight for the Uphol specs 
            psv.SetCurrentStyleSheet(ps, "monochrome.ctb");

            // We need to link the PlotInfo to the  PlotSettings and then validate it
            pi.OverrideSettings = ps;
            PlotInfoValidator piv = new PlotInfoValidator();
            piv.MediaMatchingPolicy = MatchingPolicy.MatchEnabled;
            piv.Validate(pi);

            //pass name, window, saveloc..... to plot engine
            Plot_Engine.plotEngine(pi, pdfDirectory, doc, ed, fileName);
            
        }

        //if the media size doesn't exist, this will search media list for best match
        // 8.5 x 11 should be there
        private static string setClosestMediaName(PlotSettingsValidator psv, PlotSettings ps,
            double pageWidth, double pageHeight, bool matchPrintableArea)
        {
            //get all of the media listed for plotter
            System.Collections.Specialized.StringCollection mediaList = psv.GetCanonicalMediaNameList(ps);
            double smallestOffest = 0.0;
            string selectedMedia = string.Empty;
            PlotRotation selectedRot = PlotRotation.Degrees000;

            foreach (string media in mediaList)
            {
                psv.SetCanonicalMediaName(ps, media);

                double mediaWidth = ps.PlotPaperSize.X;
                double mediaHeight = ps.PlotPaperSize.Y;

                if (matchPrintableArea)
                {
                    mediaWidth -= (ps.PlotPaperMargins.MinPoint.X + ps.PlotPaperMargins.MaxPoint.X);
                    mediaHeight -= (ps.PlotPaperMargins.MinPoint.Y + ps.PlotPaperMargins.MaxPoint.Y);
                }

                PlotRotation rot = PlotRotation.Degrees090;

                //check that we are not outside the media print area
                if (mediaWidth < pageWidth || mediaHeight < pageHeight)
                {
                    //Check if turning paper will work
                    if (mediaHeight < pageWidth || mediaWidth >= pageHeight)
                    {
                        //still too small
                        continue;
                    }
                    rot = PlotRotation.Degrees090;
                }

                double offset = Math.Abs(mediaWidth * mediaHeight - pageWidth * pageHeight);

                if (selectedMedia == string.Empty || offset < smallestOffest)
                {
                    selectedMedia = media;
                    smallestOffest = offset;
                    selectedRot = rot;

                    if (smallestOffest == 0)
                        break;
                }
            }
            psv.SetCanonicalMediaName(ps, selectedMedia);
            psv.SetPlotRotation(ps, selectedRot);
            return selectedMedia;
        }
    }
}
