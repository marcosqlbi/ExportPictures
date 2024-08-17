# ExportPictures
Export the bitmaps in a Word document by using the following description to extract the number of the picture.
The chapter number is the first number included in the file name.
The output folder will include one file for each bitmap with the format F cc ff.png where cc is the chapter number and ff is the figure number.

The tool also renumber figures in the document and the corresponding demo files.

Usage:
- **-src**:<filename> the source Word document
- **-demo**:<srcDemoFolder> the folder containing the demo files
- **-dst**:<destDemoFolder> the folder where the demo files will be copied
- **-pic**:<pictureExportFolder> the folder where the pictures will be exported
- **-newChapter**:<newChapterNumber> the new chapter number (default is the current one extracted from chapter file name)
- **-renumber** renumber the figures in the document and the demo files
- **-check** only chec the figures in the document and the demo files without saving a new Word file and without copying demo files
- **-export** export the pictures
- **-force** force the action without asking confirmation
