using System;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationToPdfConverter;
using System.IO;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Drawing;
using System.Reflection.Metadata;
using Syncfusion.Pdf.Graphics;


namespace PowerpointToPDFDemo.Controllers
{
	public class ConversionController : Controller
    {
        private readonly IWebHostEnvironment _environment;

        public ConversionController(IWebHostEnvironment environment)
        {
            _environment = environment;
        }


        [HttpGet]
        public IActionResult Index()
        {

            return View();
        }



        [HttpPost]
        public IActionResult Upload(IFormFile pptxFile)
        {

            // Generate a unique filename
            string uniqueFileName = Guid.NewGuid().ToString() + "SAMPLE.pdf";

            // Get the path to the wwwroot/uploads directory
            string uploadsPath = Path.Combine(_environment.WebRootPath, "uploads");

            // Ensure the directory exists
            Directory.CreateDirectory(uploadsPath);

            // Combine the path with the unique filename
            string filePath = Path.Combine(uploadsPath, uniqueFileName);


            //Load the PowerPoint presentation into stream.
            using (var fileStreamInput = pptxFile.OpenReadStream())
            {
                //Open the existing PowerPoint presentation with loaded stream.
                using (IPresentation pptxDoc = Presentation.Open(fileStreamInput))
                {
                    //Create the MemoryStream to save the converted PDF.
                    using (MemoryStream pdfStream = new MemoryStream())
                    {
                        //Convert the PowerPoint document to PDF document.
                        using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
                        {


                            //Adding layer
                            //PdfPage page = pdfDocument.Pages[0];
                            //// Get the page size before it is rendered.
                            //var pageSize = page.Size;
                            //// Get the width of the page.
                            //var width = pageSize.Width;
                            //PdfPageLayer layer = page.Layers.Add("Layer1");
                            //PdfGraphics graphics = layer.Graphics;
                            //graphics.TranslateTransform(0,0);
                            //PdfPen pen = new PdfPen(Syncfusion.Drawing.Color.DarkGray,50);
                            //Syncfusion.Drawing.RectangleF bounds = new Syncfusion.Drawing.RectangleF(0, 0, width, 100);
                            //graphics.DrawRectangle(pen,bounds);

                       
                            // Set the IsTrialWatermarkVisible property to false.
                            pdfDocument.PageSettings.SetMargins(0f);

                            //Get the page object.
                            pdfDocument.Save(pdfStream);
                            pdfStream.Position = 0;
                        }

                        var webRootPath = _environment.WebRootPath;

                        //Create the output PDF file stream
                        using (FileStream fileStreamOutput = new FileStream(filePath, FileMode.Create))
                        {
                            //Copy the converted PDF stream into created output PDF stream
                            pdfStream.CopyTo(fileStreamOutput);
                            pdfStream.Close();
                        }
                    }
                }
            }



            return Redirect("Index");
        }

    }
}

