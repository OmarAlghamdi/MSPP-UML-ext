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
            ClassDiagramFactory.getInstance().createClass();

        }

        private void btn_interface_Click(object sender, RibbonControlEventArgs e)
        {
            ClassDiagramFactory.getInstance().createInterface();
        }

        private void btn_associ_Click(object sender, RibbonControlEventArgs e)
        {
            ClassDiagramFactory.getInstance().createAssociation();
        }

        private void btn_aggr_Click(object sender, RibbonControlEventArgs e)
        {
            ClassDiagramFactory.getInstance().createAggregation();
        }

        private void btn_comps_Click(object sender, RibbonControlEventArgs e)
        {
            ClassDiagramFactory.getInstance().createComposition();
        }

        private void btn_gener_Click(object sender, RibbonControlEventArgs e)
        {
            ClassDiagramFactory.getInstance().createGeneralization();
        }

        private void btn_impl_Click(object sender, RibbonControlEventArgs e)
        {
            ClassDiagramFactory.getInstance().createImplementation();
        }
    }

    internal class ClassDiagramFactory
    {
        private const int ASSOCIATION = 0;
        private const int AGGREGATION = 1;
        private const int COMPOSTION = 2;
        private const int GENERALIZATION = 3;
        private const int IMPLEMENTATION = 4;

        private static ClassDiagramFactory INSTANCE = new ClassDiagramFactory();
        private int objectCount = 0; // diffrentiate object with same name

        private ClassDiagramFactory()
        {
    
        }

        public static ClassDiagramFactory getInstance()
        {
            return INSTANCE;
        }

        public void createClass()
        {
            using (ClassForm prompt = new ClassForm())
            {
                DialogResult result = prompt.ShowDialog();
                if (result == DialogResult.OK)
                {
                    createEntity(prompt.name, prompt.fields, prompt.ops);
                }
            }

        }

        public void createInterface()
        {
            using (InterfaceForm prompt = new InterfaceForm())
            {
                DialogResult result = prompt.ShowDialog();
                if (result == DialogResult.OK)
                {
                    createEntity(prompt.name, null, prompt.ops);
                }
            }
        }

        private void createEntity(String name, String[] fields, String[] ops)
        {
            Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            int heightMultiplier = 20;
            int fieldsLength = 0;

            // create name section
            Shape rectTop = slide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 0, 0, 100, heightMultiplier);
            rectTop.Fill.ForeColor.RGB = (int)XlRgbColor.rgbWheat;
            rectTop.Line.ForeColor.RGB = (int)XlRgbColor.rgbBlack;
            rectTop.Name = "rect_top" + objectCount;
            // print the name
            Shape labelName = slide.Shapes.AddLabel(
                Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, -15, 0, 0);
            labelName.TextFrame.TextRange.Text = "<<interface>>\n" + name;
            labelName.TextEffect.FontSize = 10;
            labelName.TextEffect.FontBold = Microsoft.Office.Core.MsoTriState.msoCTrue;
            labelName.Name = name + objectCount;

            // wxecute for classes only
            if (fields != null)
            {
                labelName.TextFrame.TextRange.Text = name;
                fieldsLength = fields.Length;
                // create fields section
                Shape rectMiddle = slide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 0, heightMultiplier, 100, heightMultiplier * fields.Length);
                rectMiddle.Fill.ForeColor.RGB = (int)XlRgbColor.rgbWheat;
                rectMiddle.Line.ForeColor.RGB = (int)XlRgbColor.rgbBlack;
                rectMiddle.Name = "rect_middle" + objectCount;
                // print fields
                
                for (int i = 0; i < fields.Length; i++)
                {
                    Shape labelField = slide.Shapes.AddLabel(
                        Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, heightMultiplier * (i + 1), 0, 0);
                    labelField.TextFrame.TextRange.Text = fields[i];
                    labelField.TextEffect.FontSize = 10;
                    fields[i] = fields[i] + objectCount;
                    labelField.Name = fields[i];
                }
            }

            // creates operations section
            Shape rectBottom = slide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
                0, heightMultiplier * (fieldsLength + 1), 100, heightMultiplier * ops.Length);
            rectBottom.Fill.ForeColor.RGB = (int)XlRgbColor.rgbWheat;
            rectBottom.Line.ForeColor.RGB = (int)XlRgbColor.rgbBlack;
            rectBottom.Name = "rect_bottom" + objectCount;
            // print operations
            for (int i = 0; i < ops.Length; i++)
            {
                Shape labelOp = slide.Shapes.AddLabel(
                    Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                    0, heightMultiplier * (fieldsLength + i + 1), 0, 0);
                labelOp.TextFrame.TextRange.Text = ops[i];
                labelOp.TextEffect.FontSize = 10;
                ops[i] = ops[i] + objectCount;
                labelOp.Name = ops[i];
            }


            String[] shapes;

            if (fields != null)
            {
                String[] temp = { "rect_top" + objectCount,
                "rect_middle" + objectCount,
                "rect_bottom" + objectCount, name + objectCount };
                shapes = new string[temp.Length + fields.Length + ops.Length];
                temp.CopyTo(shapes, 0);
                fields.CopyTo(shapes, temp.Length);
                ops.CopyTo(shapes, temp.Length + fields.Length);
            }
            else
            {
                String[] temp = { "rect_top" + objectCount,
                "rect_bottom" + objectCount, name + objectCount };
                shapes = new string[temp.Length + ops.Length];
                temp.CopyTo(shapes, 0);
                ops.CopyTo(shapes, temp.Length);
            }

            
            slide.Shapes.Range(shapes).Group();
            objectCount++;
        }

        

        public void createAssociation()
        {
            createRelation(ASSOCIATION);
            
        }

        public void createAggregation()
        {
            createRelation(AGGREGATION);
            
        }

        public void createComposition()
        {
            createRelation(COMPOSTION);
            
        }

        public void createGeneralization()
        {
            createRelation(GENERALIZATION);
      
        }

        public void createImplementation()
        {
            createRelation(IMPLEMENTATION);

        }

        private void createRelation(int type)
        {
            Slide Slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

            // create line
            Shape compLine;
            if (type == AGGREGATION || type == COMPOSTION)
            {
                compLine = Slide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 5, 4, 75, 0);
            }
            else
            {
                compLine = Slide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 0, 4, 75, 0);
            }
            
            compLine.Fill.ForeColor.RGB = (int)XlRgbColor.rgbBlack;
            compLine.Line.ForeColor.RGB = (int)XlRgbColor.rgbBlack;
            compLine.Line.Weight = 3;
            compLine.Name = "line" + objectCount;


            // create head of relation
            Shape compHead = null;
            
            switch (type)
            {
                case AGGREGATION:
                    compHead = Slide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeDiamond, 0, 0, 10, 10);
                    break;

                case COMPOSTION:
                    compHead = Slide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeDiamond, 0, 0, 10, 10);
                    break;

                case GENERALIZATION:
                    compHead = Slide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeIsoscelesTriangle, 73, 0, 15, 10);
                    compHead.Rotation = 90;
                    break;

                case IMPLEMENTATION:
                    compHead = Slide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeIsoscelesTriangle, 73, 0, 15, 10);
                    compHead.Rotation = 90;
                    compLine.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineRoundDot;
                    break;

            }


            if (compHead != null)
            {
                // fill the head if it is composition
                if (type == COMPOSTION)
                {
                    compHead.Fill.ForeColor.RGB = (int)XlRgbColor.rgbBlack;
                }
                else
                {
                    compHead.Fill.ForeColor.RGB = (int)XlRgbColor.rgbWhite;
                }

                compHead.Line.ForeColor.RGB = (int)XlRgbColor.rgbBlack;
                compHead.Line.Weight = 3;
                compHead.Name = "head" + objectCount;

                // groups relation object components
                String[] shapes = { "line" + objectCount, "head" + objectCount };
                Slide.Shapes.Range(shapes).Group();
            }
       
            objectCount++;
        }
    }
}
