using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;

namespace WebApplication1.Controllers
{
    [Route("api/[controller]")]
    public class ValuesController : Controller
    {
        // POST api/values
        [HttpPost]
        public void Post([FromBody]string value)
        {
        }

        // PUT api/values/5
        [HttpPut("{id}")]
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE api/values/5
        [HttpDelete("{id}")]
        public void Delete(int id)
        {
        }

        // GET api/values/FillDocument/felix/wollmux
        [HttpGet("{vorname}/{nachname}")]
        public string Get(string vorname, string nachname)
        {
            run(vorname, nachname);
            return "value";
        }

        private void run(string vorname, string nachname)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Create(@"C:\\TestDoc\test.docx", DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = doc.AddMainDocumentPart();
                var body = new Body();
                mainPart.Document = new Document(body);

                //insert content control, find and replace text in content control.
                this.InsertContentControl(body, "=fullname", "Beispiel Text");
                this.ReplaceTextInContentControlByTag(doc, "=fullname", vorname + " " + nachname);

                //create paragraph, apply custom style
                Paragraph paragraph = new Paragraph(new Run(new Text("Paragraph Style")));
                paragraph.ParagraphId = "paragraph1";

                this.CreateStyle(doc, "1", "myCustomStyle");
                this.ApplyStyleById(doc, paragraph, "1");

                //find paragraph example
                Paragraph firstParagraph =
                    body.Elements<Paragraph>().Where(w => w.ParagraphId == "paragraph1").FirstOrDefault();

                doc.Save();
            }
        }

        private void ApplyStyleById(WordprocessingDocument doc, Paragraph paragraph, string styleId)
        {
            // Add the paragraph as a child element of the w:body element.
            doc.MainDocumentPart.Document.Body.AppendChild(paragraph);

            // If the paragraph has no ParagraphProperties object, create one.
            if (paragraph.Elements<ParagraphProperties>().Count() == 0)
            {
                paragraph.PrependChild<ParagraphProperties>(new ParagraphProperties());
            }

            // Get a reference to the ParagraphProperties object.
            ParagraphProperties pPr = paragraph.ParagraphProperties;

            // If a ParagraphStyleId object doesn't exist, create one.
            if (pPr.ParagraphStyleId == null)
                pPr.ParagraphStyleId = new ParagraphStyleId();

            // Set the style of the paragraph.
            pPr.ParagraphStyleId.Val = styleId;
        }

        private WordprocessingDocument ReplaceTextInContentControlByTag(WordprocessingDocument doc, string contentControlTag, string text)
        {
            SdtElement element = doc.MainDocumentPart.Document.Body.Descendants<SdtElement>()
              .FirstOrDefault(sdt => sdt.SdtProperties.GetFirstChild<Tag>()?.Val == contentControlTag);

            if (element == null)
                throw new ArgumentException($"ContentControlTag \"{contentControlTag}\" doesn't exist.");

            element.Descendants<Text>().First().Text = text;
            element.Descendants<Text>().Skip(1).ToList().ForEach(t => t.Remove());

            return doc;
        }

        private void InsertContentControl(Body body, string tag, string contentText)
        {
            //praragraph to be added to the rich text content control
            Run run = new Run(new Text(contentText));
            Paragraph paragraph = new Paragraph(run);

            SdtProperties sdtPr = new SdtProperties(new Tag { Val = tag });
            SdtContentBlock sdtCBlock = new SdtContentBlock(paragraph);
            SdtBlock sdtBlock = new SdtBlock(sdtPr, sdtCBlock);

            body.AppendChild(sdtBlock);
        }

        private void CreateStyle(WordprocessingDocument doc, string styleId, string styleName)
        {
            // Get the Styles part for this document.
            StyleDefinitionsPart part =
                doc.MainDocumentPart.StyleDefinitionsPart;

            // If the Styles part does not exist, add it and then add the style.
            if (part == null)
            {
                part = AddStylesPartToPackage(doc);
            }

            // Access the root element of the styles part.
            Styles styles = part.Styles;
            if (styles == null)
            {
                part.Styles = new Styles();
                part.Styles.Save();
            }

            // Create a new paragraph style element and specify some of the attributes.
            Style style = new Style()
            {
                Type = StyleValues.Paragraph,
                StyleId = styleId,
                CustomStyle = true,
                Default = false
            };

            // Create and add the child elements (properties of the style).
            AutoRedefine autoredefine1 = new AutoRedefine() { Val = OnOffOnlyValues.Off };
            BasedOn basedon1 = new BasedOn() { Val = "Normal" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "OverdueAmountChar" };
            Locked locked1 = new Locked() { Val = OnOffOnlyValues.Off };
            PrimaryStyle primarystyle1 = new PrimaryStyle() { Val = OnOffOnlyValues.On };
            StyleHidden stylehidden1 = new StyleHidden() { Val = OnOffOnlyValues.Off };
            SemiHidden semihidden1 = new SemiHidden() { Val = OnOffOnlyValues.Off };
            StyleName styleName1 = new StyleName() { Val = styleName };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "Normal" };
            UIPriority uipriority1 = new UIPriority() { Val = 1 };
            UnhideWhenUsed unhidewhenused1 = new UnhideWhenUsed() { Val = OnOffOnlyValues.On };

            style.Append(autoredefine1);
            style.Append(basedon1);
            style.Append(linkedStyle1);
            style.Append(locked1);
            style.Append(primarystyle1);
            style.Append(stylehidden1);
            style.Append(semihidden1);
            style.Append(styleName1);
            style.Append(nextParagraphStyle1);
            style.Append(uipriority1);
            style.Append(unhidewhenused1);

            // Create the StyleRunProperties object and specify some of the run properties.
            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Bold bold1 = new Bold();
            Color color1 = new Color() { ThemeColor = ThemeColorValues.Hyperlink };
            RunFonts font1 = new RunFonts() { Ascii = "Lucida Console" };
            Italic italic1 = new Italic();
            // Specify a 12 point size.
            FontSize fontSize1 = new FontSize() { Val = "30" };
            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(color1);
            styleRunProperties1.Append(font1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(italic1);

            // Add the run properties to the style.
            style.Append(styleRunProperties1);

            // Add the style to the styles part.
            styles.Append(style);
        }

        // Add a StylesDefinitionsPart to the document.  Returns a reference to it.
        public static StyleDefinitionsPart AddStylesPartToPackage(WordprocessingDocument doc)
        {
            StyleDefinitionsPart part;
            part = doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            Styles root = new Styles();
            root.Save(part);
            return part;
        }
    }
}
