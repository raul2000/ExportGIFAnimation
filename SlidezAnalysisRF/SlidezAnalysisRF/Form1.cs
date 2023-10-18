
using System.Diagnostics;

namespace SlidezAnalysisRF
{
    using PowerPoint = Microsoft.Office.Interop.PowerPoint;
    using Microsoft.Office.Core;
    using Microsoft.Office.Interop.PowerPoint;

    public partial class Form1 : Form

    {

        public Form1()
        {
            InitializeComponent();
        }


        private void button1_Click(object sender, EventArgs e)
        {


        }

        private void button2_Click(object sender, EventArgs e)
        {

           


            string fileName = @"C:\rlz\movement.pptx";
            string exportName = "video_of_presentationxy";
            string exportPath = @"c:\rlz\{0}.wmv";

            Microsoft.Office.Interop.PowerPoint.Application ppApp = new Microsoft.Office.Interop.PowerPoint.Application();
            ppApp.Visible = MsoTriState.msoTrue;
            ppApp.WindowState = PpWindowState.ppWindowMinimized;
            Microsoft.Office.Interop.PowerPoint.Presentations oPresSet = ppApp.Presentations;
            Microsoft.Office.Interop.PowerPoint._Presentation oPres = oPresSet.Open(
                            fileName,
                            MsoTriState.msoFalse,
                            MsoTriState.msoFalse,
                            MsoTriState.msoFalse);



            

            try
            {
                oPres.CreateVideo(exportName);
               

                oPres.SaveCopyAs(@"c:\rlz\test123456.wmv", PpSaveAsFileType.ppSaveAsWMV);






            }
            finally
            {
                //ppApp.Quit();
            }


            }       
        }
    }
