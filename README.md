# CursoMultimetro

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PowerPointGenerator
{
    class Slide
    {
        public string Title { get; private set; }
        public string Content { get; private set; }

        public Slide(string title, string content)
        {
            Title = title;
            Content = content;
        }

        public SlidePart CreateSlidePart(PresentationPart presentationPart)
        {
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.Slide = new P.Slide(
                new CommonSlideData(
                    new ShapeTree(
                        new P.NonVisualGroupShapeProperties(
                            new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = Title },
                            new P.NonVisualGroupShapeDrawingProperties(),
                            new P.ApplicationNonVisualDrawingProperties()),
                        new P.Shape(
                            new P.NonVisualShapeProperties(
                                new P.NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Content" },
                                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
                                new P.ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                            new P.TextBody(
                                new A.BodyProperties(),
                                new A.ListStyle(),
                                new A.Paragraph(new A.Run(new A.Text(Title))),
                                new A.Paragraph(new A.Run(new A.Text(Content))))
                            )
                        )
                    )
                ));

            slidePart.Slide.Save();
            return slidePart;
        }
    }

    class PresentationCreator
    {
        public string FilePath { get; private set; }

        public PresentationCreator(string filePath)
        {
            FilePath = filePath;
        }

        public void CreatePresentation(params Slide[] slides)
        {
            using (PresentationDocument presentationDocument = PresentationDocument.Create(FilePath, PresentationDocumentType.Presentation))
            {
                PresentationPart presentationPart = presentationDocument.AddPresentationPart();
                presentationPart.Presentation = new Presentation();

                foreach (Slide slide in slides)
                {
                    slide.CreateSlidePart(presentationPart);
                }

                presentationPart.Presentation.Save();
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            string filePath = "CursoMultimetro.pptx";
            PresentationCreator creator = new PresentationCreator(filePath);

            Slide module1 = new Slide("Módulo 1: Introducción a la Electricidad y Seguridad", "Content for Module 1...");
            Slide module2 = new Slide("Módulo 2: Introducción al Multímetro Digital", "Content for Module 2...");

            creator.CreatePresentation(module1, module2);
        }
    }
}
