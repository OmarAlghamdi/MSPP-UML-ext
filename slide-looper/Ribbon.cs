using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;

namespace slide_looper
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

       
        private void sections_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.PowerPoint.Application app = Globals.ThisAddIn.Application;

            // get slide count
            int slideCount = app.ActivePresentation.Slides.Count;

            // get sections
            List<String> sections = new List<string>();
            
            
            SectionProperties sectionProperties = app.ActivePresentation.SectionProperties;
            int sectionCount = sectionProperties.Count;
            for (int i = 1; i <= sectionCount; i++)
            {
                sections.Add(sectionProperties.Name(i));
            }
       
            String msg = "There are {0} slides\nThese are the sections: \n{1}";
            if (sectionCount == 0)
            {
                msg = "There are {0} slides\nThere is not any section";
            }
            MessageBox.Show(String.Format(
                msg,
                slideCount,
                list2String(sections)));
        }

        String list2String(List<String> list)
        {
            String s = "";
            foreach(String item in list)
            {
                s += "- " + item + "\n";
            }

            return s;
        }

        private void btn_class_Click(object sender, RibbonControlEventArgs e)
        {
            using (ClassForm prompt = new ClassForm())
            {
                DialogResult result = prompt.ShowDialog();
                if (result == DialogResult.OK)
                {
                    drawClass(prompt.name, prompt.fields, prompt.ops);
                }
            }

        }

        private void btn_interface_Click(object sender, RibbonControlEventArgs e)
        {
            using (InterfaceForm prompt = new InterfaceForm()) 
            {
                DialogResult result = prompt.ShowDialog();
                if (result == DialogResult.OK)
                {
                    drawInterface(prompt.name, prompt.ops);
                }
            }
        }

        private void drawInterface(String name, String[] ops)
        {
            Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            int heightMultiplier = 20;

            Shape rectTop = slide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 0, 0, 100, heightMultiplier * 2);
            rectTop.Fill.ForeColor.RGB = (int) XlRgbColor.rgbWheat;
            rectTop.Line.ForeColor.RGB = (int) XlRgbColor.rgbBlack;
            rectTop.Name = "rect_top";
            Shape rectBottom = slide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 0, heightMultiplier * 2, 100, heightMultiplier * ops.Length);
            rectBottom.Fill.ForeColor.RGB = (int)XlRgbColor.rgbWheat;
            rectBottom.Line.ForeColor.RGB = (int)XlRgbColor.rgbBlack;
            rectBottom.Name = "rect_bottom";

            Shape labelName = slide.Shapes.AddLabel(
                Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 0, 0);
            labelName.TextFrame.TextRange.Text = "<<Interface>>\n" + name;
            labelName.TextEffect.FontSize = 10;
            labelName.TextEffect.FontBold = Microsoft.Office.Core.MsoTriState.msoCTrue;
            labelName.Name = name;
            

            for (int i = 0; i < ops.Length; i++)
            {
                Shape labelOp = slide.Shapes.AddLabel(
                Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, heightMultiplier * (i+2), 0, 0);
                labelOp.TextFrame.TextRange.Text = ops[i];
                labelOp.TextEffect.FontSize = 10;
                labelOp.Name = ops[i];
            }

            String[] temp = { "rect_top", "rect_bottom", name };
            String[] shapes = new string[temp.Length + ops.Length];
            temp.CopyTo(shapes, 0);
            ops.CopyTo(shapes, temp.Length);
            slide.Shapes.Range(shapes).Group();
        }

        private void drawClass(String name, String[] fields, String[] ops)
        {
            Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            int heightMultiplier = 20;

            Shape rectTop = slide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 0, 0, 100, heightMultiplier);
            rectTop.Fill.ForeColor.RGB = (int)XlRgbColor.rgbWheat;
            rectTop.Line.ForeColor.RGB = (int)XlRgbColor.rgbBlack;
            rectTop.Name = "rect_top";
            Shape rectMiddle = slide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 0, heightMultiplier, 100, heightMultiplier * fields.Length);
            rectMiddle.Fill.ForeColor.RGB = (int)XlRgbColor.rgbWheat;
            rectMiddle.Line.ForeColor.RGB = (int)XlRgbColor.rgbBlack;
            rectMiddle.Name = "rect_middle";
            Shape rectBottom = slide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
                0, heightMultiplier * (fields.Length + 1), 100, heightMultiplier * ops.Length);
            rectBottom.Fill.ForeColor.RGB = (int)XlRgbColor.rgbWheat;
            rectBottom.Line.ForeColor.RGB = (int)XlRgbColor.rgbBlack;
            rectBottom.Name = "rect_bottom";

            Shape labelName = slide.Shapes.AddLabel(
                Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 0, 0);
            labelName.TextFrame.TextRange.Text = name;
            labelName.TextEffect.FontSize = 10;
            labelName.TextEffect.FontBold = Microsoft.Office.Core.MsoTriState.msoCTrue;
            labelName.Name = name;

            // print fields
            for (int i = 0; i < fields.Length; i++)
            {
                Shape labelField = slide.Shapes.AddLabel(
                    Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, heightMultiplier * (i + 1), 0, 0);
                labelField.TextFrame.TextRange.Text = fields[i];
                labelField.TextEffect.FontSize = 10;
                labelField.Name = fields[i];
            }
            // print operations
            for (int i = 0; i < ops.Length; i++)
            {
                Shape labelOp = slide.Shapes.AddLabel(
                    Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                    0, heightMultiplier * (fields.Length + i + 1), 0, 0);
                labelOp.TextFrame.TextRange.Text = ops[i];
                labelOp.TextEffect.FontSize = 10;
                labelOp.Name = ops[i];
            }

            String[] temp = { "rect_top", "rect_middle", "rect_bottom", name };
            String[] shapes = new string[temp.Length+ fields.Length + ops.Length];
            temp.CopyTo(shapes, 0);
            fields.CopyTo(shapes, temp.Length);
            ops.CopyTo(shapes, temp.Length + fields.Length);
            slide.Shapes.Range(shapes).Group();
        }

        private void btn_associ_Click(object sender, RibbonControlEventArgs e)
        {
            Slide Slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

            Shape line = Slide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 0, 0, 75, 2);
            line.Fill.ForeColor.RGB = (int)XlRgbColor.rgbBlack;
            line.Line.ForeColor.RGB = (int)XlRgbColor.rgbBlack;
        }

        private void btn_aggr_Click(object sender, RibbonControlEventArgs e)
        {
            Slide Slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            // create line
            Shape compLine = Slide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 5, 4, 75, 2);
            compLine.Fill.ForeColor.RGB = (int)XlRgbColor.rgbBlack;
            compLine.Line.ForeColor.RGB = (int)XlRgbColor.rgbBlack;
            compLine.Name = "line";
            // create empty diamond
            Shape compDiamond = Slide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeDiamond, 0, 0, 10, 10);
            compDiamond.Fill.ForeColor.RGB = (int)XlRgbColor.rgbWhite;
            compDiamond.Line.ForeColor.RGB = (int)XlRgbColor.rgbBlack;
            compDiamond.Line.Weight = 3;
            compDiamond.Name = "diamond";

            String[] shapes = { "line", "diamond" };
            Slide.Shapes.Range(shapes).Group();
        }

        private void btn_comps_Click(object sender, RibbonControlEventArgs e)
        {
            Slide Slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            // create line
            Shape compLine = Slide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 5, 4, 75, 2);
            compLine.Fill.ForeColor.RGB = (int) XlRgbColor.rgbBlack;
            compLine.Line.ForeColor.RGB = (int)XlRgbColor.rgbBlack;
            compLine.Name = "line";
            // create filled diamond
            Shape compDiamond = Slide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeDiamond, 0, 0, 10, 10);
            compDiamond.Fill.ForeColor.RGB = (int) XlRgbColor.rgbBlack;
            compDiamond.Line.ForeColor.RGB = (int)XlRgbColor.rgbBlack;
            compDiamond.Name = "diamond";

            String[] shapes = { "line", "diamond" };
            Slide.Shapes.Range(shapes).Group();
        }

        private void btn_gener_Click(object sender, RibbonControlEventArgs e)
        {
            Slide Slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            // create line
            Shape compLine = Slide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 0, 4, 75, 2);
            compLine.Fill.ForeColor.RGB = (int)XlRgbColor.rgbBlack;
            compLine.Line.ForeColor.RGB = (int)XlRgbColor.rgbBlack;
            compLine.Name = "line";
            // create triangel
            Shape compTriangle = Slide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeIsoscelesTriangle, 73, 0, 15, 10);
            compTriangle.Fill.ForeColor.RGB = (int)XlRgbColor.rgbWhite;
            compTriangle.Line.ForeColor.RGB = (int)XlRgbColor.rgbBlack;
            compTriangle.Line.Weight = 3;
            compTriangle.Rotation = 90;
            compTriangle.Name = "triangel";

            String[] shapes = { "line", "triangel" };
            Slide.Shapes.Range(shapes).Group();
        }

        private void btn_impl_Click(object sender, RibbonControlEventArgs e)
        {
            Slide Slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            // create line
            Shape compLine = Slide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 0, 4, 75, 0);
            compLine.Fill.ForeColor.RGB = (int)XlRgbColor.rgbBlack;
            compLine.Line.ForeColor.RGB = (int)XlRgbColor.rgbBlack;
            compLine.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineRoundDot;
            compLine.Line.Weight = 3;
            compLine.Name = "line";
            // create triangel
            Shape compTriangle = Slide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeIsoscelesTriangle, 73, 0, 15, 10);
            compTriangle.Fill.ForeColor.RGB = (int)XlRgbColor.rgbWhite;
            compTriangle.Line.ForeColor.RGB = (int)XlRgbColor.rgbBlack;
            compTriangle.Line.Weight = 3;
            compTriangle.Rotation = 90;
            compTriangle.Name = "triangel";

            String[] shapes = { "line", "triangel" };
            Slide.Shapes.Range(shapes).Group();
        }
    }
}
