// (C) Copyright 2013 by Hydro Systems - KDI 
//
using System;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using acadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD;
using Autodesk.AutoCAD.Internal;

using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.CSharp.RuntimeBinder;


//using CarlosAg.ExcelXmlWriter;

using System.Globalization;
using System.Threading;

using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

[assembly: CommandClass(typeof(HSKDICommands.MyCommands))]

namespace HSKDICommands
{
    public class MyCommands
    {
        [CommandMethod("Extractor")]
        public void Extractor()
        {
            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            Transaction tr = db.TransactionManager.StartTransaction();
            PromptEntityResult per;

            using (tr)
            {
                List<string> layerNames = new List<string>();

                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                // Create block selector                
                                
                PromptEntityOptions peo = new PromptEntityOptions("\nSelect an object on desired Layer or <Filter by Area>.");
                peo.Keywords.Add("Filter");                
                peo.SetRejectMessage("\nObject Selection Failed. Select an object on desired Layer.");                

                PromptSelectionOptions pso = new PromptSelectionOptions();                
                pso.MessageForAdding = "\nHighlighted Objects will be included. Select individual objects or filtering area limits.";
                pso.MessageForRemoval = "\nObject removed. Select individual objects or filtering area limits.";

                //PromptPointOptions ppo = new PromptPointOptions("\nSelect first point in window or <End>. ");
                //ppo.Keywords.Add("End");

                //PromptPointResult ppr = null;
                

                PromptSelectionResult psr = null;                
                ObjectIdCollection objsInWindow = new ObjectIdCollection();
                
                do
                {                    
                    per = ed.GetEntity(peo);
                    psr = null;

                    if (per.Status == PromptStatus.Keyword && per.StringResult == "Filter")
                    {
                        peo.Message = "\nSelect an object on next desired Layer or press <Enter> when done.";
                        peo.Keywords.Clear();
                        //Point3dCollection pts = new Point3dCollection();

                        //do
                        //{
                        //    ppr = ed.GetPoint(ppo);
                        //    pts.Add(ppr.Value);
                        //    ppo.BasePoint = pts[pts.Count - 1];
                        //    ppo.UseDashedLine = true;
                        //    ppo.UseBasePoint = true;

                        //    ppo.Message = "\nSelect Next point.";
                        //    if (pts.Count >= 2) ppo.Message = "\nSelect Next point or <End> to close selection area.";
                        
                        //    if (ppr.Status == PromptStatus.OK)
                        //    {
                        //        pts.Add(ppr.Value);
                        //    }

                        //    if (ppr.Status == PromptStatus.Keyword)
                        //    {
                        //        break;
                        //    }
                            
                        //} while (ppr.Status == PromptStatus.OK && pts.Count < 3);

                        //if (pts.Contains(new Point3d(0, 0, 0))) pts.Remove(new Point3d(0, 0, 0));

                        //psr = ed.SelectWindow(pts[0], pts[3]);
                        psr = ed.GetSelection(pso);
                        SelectionSet selSet = psr.Value;                        

                        foreach (SelectedObject sel in selSet)
                        {
                            objsInWindow.Add(sel.ObjectId);
                        }
                        ed.WriteMessage("\nFilter Added.");                        
                    }

                    if (per.Status == PromptStatus.OK && per.ObjectId != null)
                    {
                        peo.Message = "\nSelect an object on next desired Layer or press <Enter> when done.";
                        peo.Keywords.Clear();                        
                        Entity ent = (Entity)tr.GetObject(per.ObjectId, OpenMode.ForRead);
                        if (ent.ObjectId != null)
                        {
                            string entType = ent.GetType().ToString().Split('.')[ent.GetType().ToString().Split('.').Length - 1];
                            if (!layerNames.Contains(ent.Layer.ToString())) layerNames.Add(ent.Layer.ToString());
                            string message = "\n\t" + entType + " added on Layer " + ent.Layer.ToString() 
                                           + ". Select an object on next desired Layer or press <Enter> when done.";
                            peo.Message = message;
                        }
                    }
                    
                } while (per.Status == PromptStatus.OK || psr != null);

                ObjectIdCollection entsOnLayers = new ObjectIdCollection();

                foreach (string layerName in layerNames)
                {
                    ObjectIdCollection entsOnLayer = HSKDICommon.Commands.GetEntitiesOnLayer(layerName);
                    foreach (ObjectId ent in entsOnLayer)
                    {
                        entsOnLayers.Add(ent);
                    }
                }
                ObjectIdCollection entsInAreaOnLayers = new ObjectIdCollection();

                if (objsInWindow.Count > 0)
                {
                    foreach(ObjectId myEnt in entsOnLayers)
                    {
                        if (objsInWindow.Contains(myEnt)) entsInAreaOnLayers.Add(myEnt);
                    }
                    entsOnLayers = entsInAreaOnLayers;
                }

                List<NewTableRow> tableRows = new List<NewTableRow>();
                foreach (ObjectId objID in entsOnLayers)
                {
                    Entity ent = (Entity)tr.GetObject(objID, OpenMode.ForRead);
                    string entType = ent.GetType().ToString().Split('.')[ent.GetType().ToString().Split('.').Length - 1];
                    double area = 0;
                    double length = 0;
                    string layer = null;
                    string blkName = null;
                    List<string> attTags = new List<string>();
                    List<string> attTexts = new List<string>();
                    TypedValue[] xData = null;
                    List<double> coords = new List<double>();
                    string txt = null;

                    switch (entType)
                    {
                        case "Polyline":
                            Polyline pl = ent as Polyline;
                            area = pl.Area;
                            length = pl.Length;
                            layer = pl.Layer.ToString();
                            //xData = pl.XData.AsArray();
                            break;
                        case "Line":
                            Autodesk.AutoCAD.DatabaseServices.Line l = ent as Autodesk.AutoCAD.DatabaseServices.Line;
                            length = l.Length;
                            layer = l.Layer.ToString();
                            //xData = l.XData.AsArray();
                            break;
                        case "BlockReference":
                            BlockReference br = ent as BlockReference;
                            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
                            blkName = btr.Name;
                            layer = br.Layer.ToString();
                            coords.Add(Math.Round(br.Position.X));
                            coords.Add(Math.Round(br.Position.Y));
                            coords.Add(Math.Round(br.Position.Z));
                            //xData = br.XData.AsArray();
                            foreach (ObjectId attId in br.AttributeCollection)
                            {
                                Entity att = (Entity)tr.GetObject(attId, OpenMode.ForRead);

                                if (att is AttributeReference)
                                {
                                    AttributeReference ar = (AttributeReference)att;
                                    attTags.Add(ar.Tag);
                                    attTexts.Add(ar.TextString);
                                } // end if
                            } // end foreach   
                            break;
                        case "DBText":
                            DBText text = ent as DBText;
                            layer = text.Layer.ToString();
                            coords.Add(Math.Round(text.Position.X));
                            coords.Add(Math.Round(text.Position.Y));
                            coords.Add(Math.Round(text.Position.Z));
                            txt = text.TextString;
                            //xData = text.XData.AsArray();
                            break;
                        case "MText":
                            MText mtext = ent as MText;
                            layer = mtext.Layer.ToString();
                            coords.Add(Math.Round(mtext.Location.X));
                            coords.Add(Math.Round(mtext.Location.Y));
                            coords.Add(Math.Round(mtext.Location.Z));
                            txt = mtext.Text;
                            //xData = mtext.XData.AsArray();
                            break;
                        default:
                            //uncatigorized. Send only type & Layer
                            layer = ent.Layer.ToString();
                            break;
                    }

                    NewTableRow tableRow = new NewTableRow(layer, entType, area, length, coords, blkName, attTags, attTexts, txt, xData);
                    int i;
                    switch (entType)
                    {
                        case "Polyline":
                            i = tableRows.FindIndex(delegate(NewTableRow t)
                                    {
                                        return t.layer == ent.Layer && t.entType == "Polyline";
                                    });
                            if (i == -1)
                            {
                                tableRows.Add(tableRow);
                            }
                            else
                            {                                
                                if (tableRow.area > 0) tableRows[i].area += tableRow.area;
                                if (tableRow.length > 0) tableRows[i].length += tableRow.length;
                            }
                            break;
                        case "Line":
                            i = tableRows.FindIndex(delegate(NewTableRow t)
                            {
                                return t.layer == ent.Layer && t.entType == "Line";
                            });
                            if (i == -1)
                            {
                                tableRows.Add(tableRow);
                            }
                            else
                            {
                                if (tableRow.length > 0) tableRows[i].length += tableRow.length;
                            }
                            break;
                        case "BlockReference":
                            tableRows.Add(tableRow);
                            break;
                        case "DBText":
                            tableRows.Add(tableRow);
                            break;
                        case "MText":
                            tableRows.Add(tableRow);
                            break;
                        default:
                            break;
                    }
                }

                if (entsOnLayers.Count > 0)
                {
                    // now that we have all the table rows, export them
                    ExportSpreadsheet(tableRows);
                }
            }
        }

        private static List<TypedValue> DBObjArrayToList(TypedValue[] arr)
        {
            List<TypedValue> lst = new List<TypedValue>();
            foreach (TypedValue obj in arr)
            {
                lst.Add(obj);
            }
            return lst;
        }
        
        public void ExportSpreadsheet(List<NewTableRow> rows) // use as template to send any data to Excel
        {
            Document myDWG;            

            myDWG = acadApp.DocumentManager.MdiActiveDocument;
            Editor ed = myDWG.Editor;

            CultureInfo oldCult = CultureInfo.CurrentCulture;
            // ' This line is very important!

            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
            //<-- change culture on whatever you need
            try
            {                   
                List<string> columnNames = new List<string>();                                           
                columnNames.Add("Layer");                
                columnNames.Add("Entity Type");
                if (rows.Exists(delegate(NewTableRow row) { return row.area > 0 ? true : false; }))
                {                    
                    columnNames.Add("Area(s)");
                }
                if (rows.Exists(delegate(NewTableRow row) { return row.length > 0 ? true : false; }))
                {
                    columnNames.Add("Length(s)");
                }
                if (rows.Exists(delegate(NewTableRow row) { return row.coords != null ? true : false; }))
                {
                    columnNames.Add("Coordinates");
                }
                if (rows.Exists(delegate(NewTableRow row) { return row.txt != null ? true : false; }))
                {
                    columnNames.Add("Text");
                }
                if (rows.Exists(delegate(NewTableRow row) { return row.blkName != null ? true : false; }))
                {
                    columnNames.Add("Block Name");
                    if (rows.Exists(delegate(NewTableRow row) { return row.attTags.Count > 0 ? true : false; }))
                    {
                        columnNames.Add("Attributes");
                    }
                }
                
                string[][] data = new string[rows.Count][];

                int i;
                for (i = 0; i < rows.Count; i++)
                {
                    List<string> row = new List<string>();
                    
                    row.Add(rows[i].layer);
                    row.Add(rows[i].entType);
                    row.Add((rows[i].area > 0) ? Math.Round(rows[i].area).ToString() : "");
                    row.Add((rows[i].length > 0) ? Math.Round(rows[i].length).ToString() : "");
                    if (rows[i].coords.Count > 0)
                        row.Add("("
                            + rows[i].coords[0].ToString() + ", "
                            + rows[i].coords[1].ToString() + ", "
                            + rows[i].coords[2].ToString() + ")");                    
                    row.Add((rows[i].txt != "") ? rows[i].txt : "");
                    row.Add((rows[i].blkName != "") ? rows[i].blkName : "");                 
                    for (int n = 0; n < rows[i].attTags.Count; n++)
                    {
                        row.Add(rows[i].attTags[n] + ":");
                        row.Add(rows[i].attTexts[n]);
                    }
                    data[i] = row.ToArray();
                }

                Excel.Application app = new Excel.Application();                
                Excel.Workbooks workbooks = (Excel.Workbooks)app.Workbooks;                
                Excel.Workbook workbook = (Excel.Workbook)(workbooks.Add(1));
                Excel.Worksheet worksheet = (Excel.Worksheet)(workbook.Sheets[1]);

                int numColumns = columnNames.Count;
                int numRows = rows.Count;
                object[,] objCol = new Object[1, numColumns];
                object[,] objData = new Object[numRows, numColumns];
               
                string columnUpper = ConvertNumberToBase26(numColumns);
                string rangeUpper = columnUpper + (numRows + 1).ToString();
                
                for(int c = 0; c < numColumns; c++)
                {
                    objCol[0, c] = columnNames[c];
                }
                Range range = worksheet.get_Range("A1", columnUpper + "1");
                range.Value = objCol;
                app.Visible = true;

                for (int r = 0; r < numRows; r++)
                {
                    for (int c = 0; c < numColumns; c++)
                    {
                        objData[r, c] = data[r][c];
                    }
                }

                range = worksheet.get_Range("A2",rangeUpper);                
                range.Value = objData;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace);
                if (ex.InnerException != null)
                {
                    MessageBox.Show("Inner exception: " + ("\n" + ex.InnerException.Message));
                }
                Thread.CurrentThread.CurrentCulture = oldCult;
            }
        }

        private static string ConvertNumberToBase26(int numColumns)
        {
            string columnAlpha = null;
            int alpha = numColumns / 27;
            int remainder = numColumns - (alpha * 26);
            if (alpha > 0) columnAlpha = (Convert.ToChar(alpha + 64)).ToString();
            if (remainder > 0) columnAlpha += (Convert.ToChar(remainder + 64)).ToString();
            return columnAlpha;
        }          
    }



    public class NewTableRow
    {
        public string layer = null;
        public string entType = null;
        public double area;
        public double length;
        public List<double> coords = new List<double>();
        public string blkName = null;
        public List<string> attTags = new List<string>();
        public List<string> attTexts = new List<string>();
        public TypedValue[] xData = null;
        public string txt = null;

        public NewTableRow(string layer, string entType, double area, double length, List<double> coords, string blkName, List<string> attTags, List<string> attTexts, string txt, TypedValue[] xData)
        {
            this.entType = entType;
            this.area = area;
            this.length = length;
            this.coords = coords;
            this.layer = layer;
            this.blkName = blkName;
            this.attTags = attTags;
            this.attTexts = attTexts;
            this.txt = txt;
            this.xData = xData;
        }
    }
}
