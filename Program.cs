using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Linq;
using System.IO.Packaging;
using System.Text;
using System.Xml.Linq;
using System.Xml;
using WebBackLibrary.Service;


List<(int oldChapterNumber, int newChapterNumber, int oldFigureNumber, int newFigureNumber)> GetUnreferencedFigures(WordprocessingDocument wordDoc, List<(int oldChapterNumber, int newChapterNumber, int oldFigureNumber, int newFigureNumber)> figures)
{
    var paragraphs = wordDoc.MainDocumentPart.Document.Body.Descendants<Paragraph>();
    var regex = new Regex(@"\s(\d{1,2})-(\d+)");
    var referencedFigures = new HashSet<(int chapterNumber, int figureNumber)>();

    foreach (var paragraph in paragraphs)
    {
        if (!IsParagraphOfStyle(paragraph, "Fig-Graphic") && !IsParagraphOfStyle(paragraph, "Num-Caption"))
        {
            var paragraphText = string.Concat(paragraph.Descendants<Text>().Select(t => t.Text));
            var matches = regex.Matches(paragraphText);
            foreach (Match match in matches)
            {
                if (match.Success)
                {
                    int chapterReference = int.Parse(match.Groups[1].Value);
                    int figureReference = int.Parse(match.Groups[2].Value);
                    referencedFigures.Add((chapterReference, figureReference));
                }
            }
        }
    }

    var unreferencedFigures = figures.Where(figure =>
        !referencedFigures.Contains((figure.oldChapterNumber, figure.oldFigureNumber))).ToList();

    return unreferencedFigures;
}

bool IsParagraphOfStyle(Paragraph paragraph, string styleId)
{
    return paragraph.ParagraphProperties != null &&
           paragraph.ParagraphProperties.ParagraphStyleId != null &&
           paragraph.ParagraphProperties.ParagraphStyleId.Val?.Value == styleId;
}

void DumpUnreferencedFigures(List<(int oldChapterNumber, int cìnewChapterNumber, int oldFigureNumber, int newFigureNumber)> unreferencedFigures)
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("UNREFERENCED FIGURES in Word Document");
    Console.ForegroundColor = ConsoleColor.White;
    foreach (var unreferencedFigure in unreferencedFigures)
    {
        Console.WriteLine($"{unreferencedFigure.oldChapterNumber}-{unreferencedFigure.oldFigureNumber}");
    }
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("--------------------");
    Console.ResetColor();
}

void ReplaceFigureReferences(WordprocessingDocument wordDoc, List<(int oldChapterNumber, int newChapterNumber, int oldFigureNumber, int newFigureNumber)> figures)
{
    figures = figures.OrderByDescending(f => f.oldFigureNumber).ToList();
    foreach (var f in figures)
    {
        var replaceWhat = $" {f.oldChapterNumber}-{f.oldFigureNumber}";
        var replaceFor = $"*{f.oldChapterNumber}*{f.oldFigureNumber}";
        // Console.WriteLine($"Replacing {replaceWhat} with {replaceFor}");
        WordDocumentService.ReplaceStringInWordDocument(wordDoc, replaceWhat.ToString(), replaceFor);
    }

    foreach (var f in figures)
    {
        var replaceWhat = $"*{f.oldChapterNumber}*{f.oldFigureNumber}";
        var replaceFor = $" {f.newChapterNumber}-{f.newFigureNumber}";
        Console.WriteLine($"Replacing {f.oldChapterNumber}-{f.oldFigureNumber} with {replaceFor}");
        WordDocumentService.ReplaceStringInWordDocument(wordDoc, replaceWhat.ToString(), replaceFor);
    }
}

List<string> GetOrphanFigureReferences(WordprocessingDocument wordDoc, List<(int oldChapterNumber, int newChapterNumber, int oldFigureNumber, int newFigureNumber)> figures)
{
    var paragraphs = wordDoc.MainDocumentPart.Document.Body.Descendants<Paragraph>();
    var regex = new Regex(@"\s(\d{1,2})-(\d+)");
    var orphanFigureReferences = new List<string>();

    foreach (var paragraph in paragraphs)
    {
        foreach (var text in paragraph.Descendants<Text>())
        {
            var matches = regex.Matches(text.Text);
            foreach (Match match in matches)
            {
                if (match.Success)
                {
                    var chapterReference = int.Parse(match.Groups[1].Value);
                    var figureReference = int.Parse(match.Groups[2].Value);

                    // Check if the figure reference exists in the figures list
                    bool existsInFigures = figures.Any(figure =>
                        figure.oldChapterNumber == chapterReference &&
                        figure.oldFigureNumber == figureReference);

                    if (!existsInFigures)
                    {
                        orphanFigureReferences.Add($"{match.Value} --> {paragraph.InnerText}");
                    }
                }
            }
        }
    }
    return orphanFigureReferences;
}


IEnumerable<(WordprocessingDocument wordDoc, Paragraph paragraph, Paragraph nextParagraph)> IterateNumCaptionParagraphs(WordprocessingDocument wordDoc)
{
    bool IsParagraphOfStyle(Paragraph paragraph, string styleId)
    {
        return paragraph.ParagraphProperties != null &&
               paragraph.ParagraphProperties.ParagraphStyleId != null &&
               paragraph.ParagraphProperties.ParagraphStyleId?.Val?.Value == styleId;
    }

    int GetParagraphIndex(Paragraph paragraph)
    {
        var body = paragraph.Parent;
        if (body != null)
        {
            var paragraphs = body.Elements<Paragraph>().ToList();
            return paragraphs.IndexOf(paragraph);
        }
        return -1;
    }


    var paragraphs = wordDoc.MainDocumentPart?.Document.Body?.Descendants<Paragraph>()
        .Where(p => IsParagraphOfStyle(p, "Fig-Graphic") || IsParagraphOfStyle(p, "SbarFig-Graphic"))
        .OrderBy(p => GetParagraphIndex(p));
    if (paragraphs != null)
    {
        foreach (var paragraph in paragraphs)
        {
            var nextParagraph = paragraph.NextSibling<Paragraph>();
            if (nextParagraph != null)
            {
                if (IsParagraphOfStyle(paragraph, "Fig-Graphic"))
                {
                    if (IsParagraphOfStyle(nextParagraph, "Num-Caption"))
                    {
                        yield return (wordDoc, paragraph, nextParagraph);
                    }
                    else
                    {
                        throw new InvalidOperationException($"Stile incorrect - there should be Num-Caption: {nextParagraph?.InnerText}");
                    }
                } 
                else if (IsParagraphOfStyle(paragraph, "SbarFig-Graphic"))
                {
                    if (IsParagraphOfStyle(nextParagraph, "SbarNum-Caption"))
                    {
                        yield return (wordDoc, paragraph, nextParagraph);
                    }
                    else
                    {
                        throw new InvalidOperationException($"Stile incorrect - there should be Sbar Num-Caption: {nextParagraph?.InnerText}");
                    }
                }
                else
                {
                    throw new InvalidOperationException($"Stile incorrect - there should be Fig-Graphic or Sbar Fig-Graphic: {paragraph?.InnerText}");
                }
            }
        }
    }
}

List<(int oldChapterNumber, int newChapterNumber, int oldFigureNumber, int newFigureNumber)> ExtractFigureList(WordprocessingDocument wordDoc, int destChapterNumber)
{
    List<(int oldChapterNumber, int newChapterNumber, int oldFigureNumber, int newFigureNumber)> figures = new List<(int, int, int, int)>();

    int index = 0;
    foreach (var p in IterateNumCaptionParagraphs(wordDoc))
    {
        string imageName = ExtractTextWithFontStyle(p.nextParagraph, "FigNum");
        if (!string.IsNullOrWhiteSpace(imageName))
        {
            var figureString = ParseFigureString(imageName);
            if (figures.Count(f => f.oldChapterNumber == figureString.chapterNumber && f.oldFigureNumber == figureString.figureNumber) > 0)
            {
                throw new InvalidOperationException($"Duplicate figure number {figureString.chapterNumber}-{figureString.figureNumber}.");
            }
            figures.Add((figureString.chapterNumber, destChapterNumber, figureString.figureNumber, ++index));
        }
    }
    return figures;
}

void CopyDemoFiles(List<(int oldChapterNumber, int newChapterNumber, int oldFigureNumber, int newFigureNumber)> figures, string sourceFolder, string destinationFolder, int sourceChapterNumber)
{
    Console.WriteLine($"Source folder: {sourceFolder}");
    Console.WriteLine($"Dest. folder:  {destinationFolder}");
    foreach (var figure in figures)
    {
        string oldFileName = $"F {figure.oldChapterNumber:D2} {figure.oldFigureNumber:D2}.*";
        string[] matchingFiles = Directory.GetFiles(sourceFolder, oldFileName);

        foreach (var matchingFile in matchingFiles)
        {
            string newFileName = $"F {figure.newChapterNumber:D2} {figure.newFigureNumber:D2}" + Path.GetExtension(matchingFile);
            string destinationPath = Path.Combine(destinationFolder, newFileName);
            File.Copy(matchingFile, destinationPath,overwrite:true);
            Console.WriteLine($"Copied: {Path.GetFileName(matchingFile)} to {Path.GetFileName(destinationPath)}");
        }
    }
}
List<(int oldChapterNumber, int newChapterNumber, int oldFigureNumber, int newFigureNumber)> GetOrphanFigures(List<(int oldChapterNumber, int newChapterNumber, int oldFigureNumber, int newFigureNumber)> figures, string sourceFolder)
{
    List<(int oldChapterNumber, int newChapterNumber, int oldFigureNumber, int newFigureNumber)> orphanFigures = new List<(int, int, int, int)>();

    foreach (var figure in figures)
    {
        string oldFileName = $"F {figure.oldChapterNumber:D2} {figure.oldFigureNumber:D2}.*";
        string[] matchingFiles = Directory.GetFiles(sourceFolder, oldFileName);

        if (matchingFiles.Length == 0)
        {
            orphanFigures.Add(figure);
        }
    }
    return orphanFigures;
}

List<string> GetOrphanFiles(List<(int oldChapterNumber, int chapterNumber, int oldFigureNumber, int newFigureNumber)> figures, string sourceFolder)
{
    List<string> orphanFiles = new List<string>();

    string[] sourceFiles = Directory.GetFiles(sourceFolder);

    foreach (var sourceFile in sourceFiles)
    {
        string fileName = Path.GetFileName(sourceFile);

        bool existsInFigures = figures.Any(figure =>
            fileName.StartsWith($"F {figure.oldChapterNumber:D2} {figure.oldFigureNumber:D2}"));

        if (!existsInFigures)
        {
            orphanFiles.Add(fileName);
        }
    }
    return orphanFiles;
}

void DumpOrphanFiles(List<string> orphanFiles)
{
    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.WriteLine("ORPHAN FILES");
    Console.ForegroundColor = ConsoleColor.White;
    foreach (var orphanFile in orphanFiles)
    {
        Console.WriteLine(orphanFile);
    }
    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.WriteLine("------------");
    Console.ResetColor();
}

void DumpOrphanFigures(List<(int oldChapterNumber, int newChapterNumber, int oldFigureNumber, int newFigureNumber)> orphanFigures)
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("ORPHAN FIGURES");
    Console.ForegroundColor = ConsoleColor.White;
    foreach (var orphanFigure in orphanFigures)
    {
        Console.WriteLine($"F {orphanFigure.oldChapterNumber:D2} {orphanFigure.oldFigureNumber:D2}.* was not found in the source folder.");
    }
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("--------------");
    Console.ResetColor();
}


void DumpOrphanFigureReferences(List<string> orphanFigureReferences)
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("ORPHAN FIGURE REFERENCES in Word Document");
    Console.ForegroundColor = ConsoleColor.White;
    foreach (var orphanFigureReference in orphanFigureReferences)
    {
        Console.WriteLine(orphanFigureReference);
    }
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("--------------");
    Console.ResetColor();
}

void ExportImagesWithSpecificStyleFromWordDocument(WordprocessingDocument wordDoc, int chapterNumber, string exportFolder)
{

    foreach (var p in IterateNumCaptionParagraphs(wordDoc))
    {
        // Extract text with "Fig Num" font style
        string imageName = ExtractTextWithFontStyle(p.nextParagraph, "FigNum");
        if (!string.IsNullOrWhiteSpace(imageName))
        {
            string imageFileName = MakeValidFileName(imageName, ref chapterNumber);

            foreach (var drawing in p.paragraph.Descendants<Drawing>())
            {
                DocumentFormat.OpenXml.Drawing.Blip blip = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
                if (blip != null)
                {
                    string embed = blip.Embed;
                    ImagePart imagePart = (ImagePart)(p.wordDoc).MainDocumentPart.GetPartById(embed);

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

string ExtractTextWithFontStyle(Paragraph paragraph, string fontStyleId)
{
    var runsWithStyle = paragraph.Descendants<Run>()
                                 .Where(r => r.RunProperties != null &&
                                             r.RunProperties.RunStyle != null &&
                                             r.RunProperties.RunStyle.Val.Value == fontStyleId && 
                                             r.RsidRunDeletion == null
                                        );

    return string.Concat(runsWithStyle.Select(r => r.InnerText));
}

// Parse string in the format "Figure 1-23" where 1 is the chapter number and 23 is the figure number
(int chapterNumber, int figureNumber) ParseFigureString(string input)
{
    const int MaxNumber = 99;
    const string ExpectedPrefix = "Figure";

    if (string.IsNullOrWhiteSpace(input))
    {
        throw new ArgumentException($"Input string \"{input}\" cannot be null or empty.");
    }

    if (!input.StartsWith(ExpectedPrefix))
    {
        throw new ArgumentException($"Invalid format \"{input}\". The name must start with 'Figure'.");
    }

    var match = Regex.Match(input, @"Figure\s+(\d+)-(\d+)");
    if (!match.Success)
    {
        throw new ArgumentException($"Invalid format \"{input}\". Expected 'Figure X-Y'.");
    }

    if (!int.TryParse(match.Groups[1].Value, out int chapterNumber) || chapterNumber < 1 || chapterNumber > MaxNumber)
    {
        throw new ArgumentException($"Invalid chapter number \"{input}\". It must be a valid number between 1 and 99.");
    }

    if (!int.TryParse(match.Groups[2].Value, out int figureNumber) || figureNumber < 1 || figureNumber > MaxNumber)
    {
        throw new ArgumentException($"Invalid figure number \"{input}\". It must be a valid number between 1 and 99.");
    }

    return (chapterNumber, figureNumber);
}


string MakeValidFileName(string name, ref int chapterNumber)
{
    const int MaxImages = 99; // Maximum number of images
    const string ExpectedPrefix = "Figure";

    // Validate the name prefix
    if (!name.StartsWith(ExpectedPrefix))
    {
        throw new ArgumentException("Invalid format. The name must start with 'Figure'.");
    }

    var cf = ParseFigureString(name);
    if (cf.chapterNumber != chapterNumber)
    {
        throw new ArgumentException("The chapter number in the file name does not match the expected chapter number.");
    }
    chapterNumber = cf.chapterNumber;
    int figureNumber = cf.figureNumber;

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
    Match matchFilename = regex.Match(Path.GetFileName(filePath));
    Match matchPath = regex.Match(filePath);


    if (matchFilename.Success)
    {
        return int.Parse(matchFilename.Value); // Convert the found number to an integer
    }
    else if (matchPath.Success)
    {
        return int.Parse(matchPath.Value); // Convert the found number to an integer
    }
    else
    {
        throw new ArgumentException("No chapter number found in the file path.");
    }
}

string ReplaceChapterNumber(string filePath, int newChapterNumber)
{
    string newFilePath = Regex.Replace(filePath, @"\d+", $"{newChapterNumber:D2}");
    return newFilePath;
}

// Test renumber
// -src:"C:\temp\src\03 - Introducing the filter context and CALCULATE.docx" -demo:"C:\temp\src" -dst:"C:\Temp\tmp" -pic:"C:\Temp\tmppic" -newChapter:4  -renumber -export 

bool renumber = false;
bool export = false;
bool checkOnly = false;
string srcFile = null;
string demoFolder = null;
string dstFolder = null;
string picFolder = null;
int? newChapterNumber = null;
bool force = false;

foreach (var arg in args)
{
    if (arg.Equals("-renumber", StringComparison.OrdinalIgnoreCase))
    {
        renumber = true;
    }
    else if (arg.Equals("-check", StringComparison.OrdinalIgnoreCase))
    {
        checkOnly = true;
    }
    else if (arg.Equals("-export", StringComparison.OrdinalIgnoreCase))
    {
        export = true;
    }
    else if (arg.StartsWith("-src:", StringComparison.OrdinalIgnoreCase))
    {
        srcFile = arg.Substring(5);
    }
    else if (arg.StartsWith("-demo:", StringComparison.OrdinalIgnoreCase))
    {
        demoFolder = arg.Substring(6);
    }
    else if (arg.StartsWith("-dst:", StringComparison.OrdinalIgnoreCase))
    {
        dstFolder = arg.Substring(5);
    }
    else if (arg.StartsWith("-pic:", StringComparison.OrdinalIgnoreCase))
    {
        picFolder = arg.Substring(5);
    }
    else if (arg.StartsWith("-newChapter:", StringComparison.OrdinalIgnoreCase))
    {
        if (int.TryParse(arg.Substring(12), out int chapter))
        {
            newChapterNumber = chapter;
        }
        else
        {
            Console.WriteLine("Invalid new chapter number.");
            return;
        }
    }
    else if (arg.Equals("-force", StringComparison.OrdinalIgnoreCase))
    {
        force = true;
    }
}

if (string.IsNullOrWhiteSpace(srcFile)
    || (string.IsNullOrWhiteSpace(demoFolder) && renumber)
    || (string.IsNullOrWhiteSpace(dstFolder) && renumber)
    || (string.IsNullOrWhiteSpace(picFolder) && export)
    )
{
    Console.ForegroundColor = ConsoleColor.Yellow;

    if (string.IsNullOrWhiteSpace(demoFolder) && renumber)
    {
        Console.WriteLine("Demo folder is required for renumbering.");
    }
    if (string.IsNullOrWhiteSpace(dstFolder) && renumber)
    {
        Console.WriteLine("Destination folder is required for renumbering.");
    }
    if (string.IsNullOrWhiteSpace(picFolder) && export)
    {
        Console.WriteLine("Picture folder is required for exporting images.");
    }

    Console.ResetColor();
    Console.WriteLine("Usage: -src:<filename> [-demo:<srcDemoFolder>] [-dst:<destDemoFolder>] [-pic:<pictureExportFolder>] [-newChapter:<newChapterNumber>] [-renumber] [-check] [-force] [-export]");
    return;
}

// Proceed with the rest of the program logic using the parsed parameters
Console.WriteLine($"Renumber  : {renumber}");
Console.WriteLine($"Check only: {checkOnly}");
Console.WriteLine($"Export Pic: {export}");
Console.WriteLine($"Force     : {force}");
Console.WriteLine($"Source File       : {srcFile}");
Console.WriteLine($"Demo Folder       : {demoFolder}");
Console.WriteLine($"Destination Folder: {dstFolder}");
Console.WriteLine($"Picture Folder    : {picFolder}");
Console.WriteLine($"New Chapter Number: {newChapterNumber}");

if (!force)
{
    Console.WriteLine("Do you want to proceed? (Y/N)");
    string response = Console.ReadLine();
    if (!response.Equals("Y", StringComparison.OrdinalIgnoreCase))
    {
        return;
    }
}

try
{
    var sourceChapterNumber = ExtractChapterNumber(srcFile);

    // Use original chapter number if a new chapter number is not required
    var destChapterNumber = newChapterNumber ?? sourceChapterNumber;

    // Copy the wordDoc to a different file
    string fileName = Path.GetFileName(srcFile);
    if (sourceChapterNumber != destChapterNumber) fileName = ReplaceChapterNumber(fileName, destChapterNumber);
    var destFilePath = Path.Combine(dstFolder, fileName);

    File.Copy(srcFile, destFilePath, overwrite: true);
    srcFile = destFilePath;
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(srcFile, isEditable: renumber))
    {
        if (checkOnly || renumber)
        {
            var figures = ExtractFigureList(wordDoc, sourceChapterNumber);

            var orphanFiles = GetOrphanFiles(figures, demoFolder);
            if (orphanFiles.Count > 0) DumpOrphanFiles(orphanFiles);
            var orphanFigures = GetOrphanFigures(figures, demoFolder);
            if (orphanFigures.Count > 0) DumpOrphanFigures(orphanFigures);
            var orphanFigureReferences = GetOrphanFigureReferences(wordDoc, figures);
            if (orphanFigureReferences.Count > 0) DumpOrphanFigureReferences(orphanFigureReferences);
            var unreferencedFigures = GetUnreferencedFigures(wordDoc, figures);
            if (unreferencedFigures.Count > 0) DumpUnreferencedFigures(unreferencedFigures);

            if (orphanFigures.Count + orphanFigureReferences.Count + unreferencedFigures.Count > 0)
            {
                throw new InvalidOperationException("There are orphan figures or references in the document.");
            }
            if (renumber)
            {
                ReplaceFigureReferences(wordDoc, figures);
                if (!checkOnly) 
                {
                    wordDoc.Save();
                    CopyDemoFiles(figures, demoFolder, dstFolder, sourceChapterNumber);
                }
            }
        }

        // Export images from the document if required
        if (!checkOnly && export)
        {
            ExportImagesWithSpecificStyleFromWordDocument(wordDoc, destChapterNumber, picFolder);
        }

    }
}
catch (InvalidOperationException ex)
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("OPERATION FAILED");
    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.WriteLine(ex.Message);
    Console.ResetColor();
}

