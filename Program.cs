using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using System.Text.RegularExpressions;

void ExportImagesWithSpecificStyleFromWordDocument(string filePath, string exportFolder)
{
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
    {
        int chapterNumber = ExtractChapterNumber(filePath);
        var paragraphs = wordDoc.MainDocumentPart.Document.Body.Descendants<Paragraph>();

        foreach (var paragraph in paragraphs)
        {
            if (IsParagraphOfStyle(paragraph, "Fig-Graphic"))
            {
                var nextParagraph = paragraph.NextSibling<Paragraph>();
                if (nextParagraph != null && IsParagraphOfStyle(nextParagraph, "Num-Caption"))
                {
                    // Extract text with "Fig Num" font style
                    string imageName = ExtractTextWithFontStyle(nextParagraph, "FigNum");
                    if (!string.IsNullOrWhiteSpace(imageName))
                    {
                        string imageFileName = MakeValidFileName(imageName, chapterNumber);

                        foreach (var drawing in paragraph.Descendants<Drawing>())
                        {
                            DocumentFormat.OpenXml.Drawing.Blip blip = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
                            if (blip != null)
                            {
                                string embed = blip.Embed;
                                ImagePart imagePart = (ImagePart)wordDoc.MainDocumentPart.GetPartById(embed);

                                string extension = imagePart.ContentType == "image/png" ? ".png" : ".jpg";
                                string fullPath = Path.Combine(exportFolder, imageFileName + extension);

                                using (Stream stream = imagePart.GetStream())
                                using (FileStream fileStream = new FileStream(fullPath, FileMode.Create))
                                {
                                    stream.CopyTo(fileStream);
                                    Console.WriteLine($"Exported: {fullPath}");
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}

string ExtractTextWithFontStyle(Paragraph paragraph, string fontStyleId)
{
    var runsWithStyle = paragraph.Descendants<Run>()
                                 .Where(r => r.RunProperties != null &&
                                             r.RunProperties.RunStyle != null &&
                                             r.RunProperties.RunStyle.Val.Value == fontStyleId);

    return string.Concat(runsWithStyle.Select(r => r.InnerText));
}

bool IsParagraphOfStyle(Paragraph paragraph, string styleId)
{
    return paragraph.ParagraphProperties != null &&
           paragraph.ParagraphProperties.ParagraphStyleId != null &&
           paragraph.ParagraphProperties.ParagraphStyleId.Val.Value == styleId;
}

string MakeValidFileName(string name, int chapterNumber)
{
    const int MaxImages = 99; // Maximum number of images
    const string ExpectedPrefix = "Figure";

    // Validate the name prefix
    if (!name.StartsWith(ExpectedPrefix))
    {
        throw new ArgumentException("Invalid format. The name must start with 'Figure'.");
    }

    // Extract the figure number from the name
    if (!int.TryParse(name.Substring(ExpectedPrefix.Length).Trim(), out int figureNumber))
    {
        throw new ArgumentException("Invalid format. Unable to parse figure number.");
    }

    // Check if figure number exceeds the limit
    if (figureNumber > MaxImages)
    {
        throw new InvalidOperationException($"The number of images exceeds the maximum limit of {MaxImages}.");
    }

    // Format the chapter and figure numbers
    string formattedFileName = $"F {chapterNumber:D2} {figureNumber:D2}";

    return formattedFileName;
}
int ExtractChapterNumber(string filePath)
{
    Regex regex = new Regex(@"\d+"); // Regular expression to find one or more digits
    Match match = regex.Match(filePath);

    if (match.Success)
    {
        return int.Parse(match.Value); // Convert the found number to an integer
    }
    else
    {
        throw new ArgumentException("No chapter number found in the file path.");
    }
}

// See https://aka.ms/new-console-template for more information
if (args.Length < 2)
{
    Console.WriteLine("Usage: ExportPictures <filePath> <exportFolder>");
    return;
}
string filePath = args[0];
string exportFolder = args[1];

if (string.IsNullOrWhiteSpace(filePath) || string.IsNullOrWhiteSpace(exportFolder))
{
    Console.WriteLine("Usage: ExportPictures <filePath> <exportFolder>");
    return;
}

//string filePath = @"C:\Doc\Study\VBA Word\Demo 05 Patterns.docx"; // Path to the Word file
//string exportFolder = @"C:\Temp\tmp";  // Folder to export images

ExportImagesWithSpecificStyleFromWordDocument(filePath, exportFolder);


    
    
