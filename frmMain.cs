using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml;
using System.IO;
using iTextSharp.text.pdf.parser;
using iTextSharp.text.pdf;

namespace Word_To_PDF
{
    public partial class frmMain : Form
    {
        private bool isDragging = false;
        private int offsetX, offsetY;
        // Create an OpenFileDialog for selecting a PDF file
        OpenFileDialog pdfFileDialog = new OpenFileDialog
        {
            Filter = "PDF Files|*.pdf",
            Title = "Select a PDF File"
        };


        public frmMain(int patientId = 0)
        {
            InitializeComponent();
            Load += form_Load;
            panel1.MouseDown += Panel1_MouseDown;
            panel1.MouseMove += Panel1_MouseMove;
            panel1.MouseUp += Panel1_MouseUp;
            panel2.MouseDown += Panel1_MouseDown;
            panel2.MouseMove += Panel1_MouseMove;
            panel2.MouseUp += Panel1_MouseUp;
        }

        private void form_Load(object sender, EventArgs e)
        {
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {

            if (pdfFileDialog.ShowDialog() == DialogResult.OK)
                    lblFileName.Text = pdfFileDialog.FileName;
        }

        private void bbtnAddTreatment_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(pdfFileDialog.FileName))
            {
                MessageBox.Show("Select a file!"); return;
            }
            // Create a SaveFileDialog for choosing the Word document save location
            SaveFileDialog wordSaveFileDialog = new SaveFileDialog
            {
                Filter = "Word Documents|*.docx",
                Title = "Save Word Document"
            };
            if (wordSaveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string wordSaveFilePath = wordSaveFileDialog.FileName;

                // Now, you have the paths to the selected PDF file and Word document save location
                // You can proceed to convert the PDF to Word and save it to the specified location
                ConvertPdfToWord(pdfFileDialog.FileName, wordSaveFilePath);
                MessageBox.Show("Completed!");
            }
        }

        private void ConvertPdfToWord(string fileName, string wordSaveFilePath)
        {
            // Open the PDF file
            using (PdfReader pdfReader = new PdfReader(fileName))
            {
                string pdfText = "";
                for (int page = 1; page <= pdfReader.NumberOfPages; page++)
                {
                    //ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                    //pdfText += PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy);

                    ITextExtractionStrategy strategy = new LocationTextExtractionStrategy();
                    pdfText += PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy);


                    PdfDictionary pageDict = pdfReader.GetPageN(page);
                    PdfDictionary resources = pageDict.GetAsDict(PdfName.RESOURCES);
                    ExtractImagesFromPage(resources, pdfReader);
                }

                // Create a new Word document
                using (var wordprocessingDocument =
                    WordprocessingDocument.Create(wordSaveFilePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = wordprocessingDocument.AddMainDocumentPart();
                    mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
                    Body body = mainPart.Document.AppendChild(new Body());
                    DocumentFormat.OpenXml.Wordprocessing.Paragraph para = body.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());
                    DocumentFormat.OpenXml.Wordprocessing.Run run = para.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Run());
                    run.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Text(pdfText));
                    wordprocessingDocument.MainDocumentPart.Document.Save();
                }
            }
        }

        // Function to extract images
        void ExtractImagesFromPage(PdfDictionary resources, PdfReader pdfReader)
        {
            if (resources == null) return;

            PdfDictionary xObjects = resources.GetAsDict(PdfName.XOBJECT);
            if (xObjects == null) return;

            foreach (PdfName name in xObjects.Keys)
            {
                PdfObject obj = xObjects.Get(name);
                if (obj != null && obj.IsIndirect())
                {
                    PdfDictionary imgObject = (PdfDictionary)PdfReader.GetPdfObject(obj);

                    if (imgObject != null && imgObject.Get(PdfName.SUBTYPE) != null && imgObject.Get(PdfName.SUBTYPE).Equals(PdfName.IMAGE))
                    {
                        int xrefIdx = Convert.ToInt32(((PRIndirectReference)obj).Number.ToString(System.Globalization.CultureInfo.InvariantCulture));
                        PdfObject pdfObj = pdfReader.GetPdfObject(xrefIdx);
                        PdfStream pdfStrem = (PdfStream)pdfObj;
                        byte[] bytes = PdfReader.GetStreamBytesRaw((PRStream)pdfStrem);

                        // 'bytes' now contains the bytes of the image; you can save/process it as needed
                        // Example: File.WriteAllBytes("outputImage.png", bytes);
                    }
                }
            }
        }

        public static void InsertImage(MainDocumentPart mainPart, string imagePath)
        {
            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
            //using (FileStream stream = new FileStream(imagePath, FileMode.Open))
            //{
            //    imagePart.FeedData(stream);
            //}

            string imageRelationshipId = mainPart.GetIdOfPart(imagePart);

            // Define the image dimensions
            long cx = 3000000L; // Width in EMUs
            long cy = 2000000L; // Height in EMUs

            // Create a drawing element with the image
            Drawing drawing = new Drawing();
            DW.Inline inline = new DW.Inline(
                new DW.Extent() { Cx = cx, Cy = cy },
                new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new DW.DocProperties() { Id = (UInt32Value)1U, Name = "Picture 1" },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks() { NoChangeAspect = true }),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "New Bitmap Image.jpg" },
                                new PIC.NonVisualPictureDrawingProperties()
                            ),
                            new PIC.BlipFill(
                                new A.Blip() { Embed = imageRelationshipId, CompressionState = A.BlipCompressionValues.Print },
                                new A.Stretch(new A.FillRectangle())
                            )
                        )
                    )
                )
            );
            inline.Append(drawing);

            // Add the image to the paragraph
            A.Paragraph imageParagraph = new A.Paragraph(new A.Run(inline));
            mainPart.Document.Body.Append(imageParagraph);
        }

        private void Panel1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isDragging = true;
                offsetX = e.X;
                offsetY = e.Y;
            }
        }

        private void Panel1_MouseMove(object sender, MouseEventArgs e)
        {
            if (isDragging)
            {
                this.Left = e.X + this.Left - offsetX;
                this.Top = e.Y + this.Top - offsetY;
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void Panel1_MouseUp(object sender, MouseEventArgs e)
        {
            isDragging = false;
        }
    }
}
