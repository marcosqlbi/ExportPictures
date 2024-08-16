using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WebBackLibrary.Service
{
    // Original code from Dmitrii Matunin
    // https://stackoverflow.com/questions/18316873/replace-text-in-word-document-using-open-xml
    public static class WordDocumentService
    {
        private class WordMatchedPhrase
        {
            public int charStartInFirstPar { get; set; }
            public int charEndInLastPar { get; set; }

            public int firstCharParOccurance { get; set; }
            public int lastCharParOccurance { get; set; }
        }

        public static WordprocessingDocument ReplaceStringInWordDocument(WordprocessingDocument wordprocessingDocument, string replaceWhat, string replaceFor)
        {
            List<WordMatchedPhrase> matchedPhrases = FindWordMatchedPhrases(wordprocessingDocument, replaceWhat);

            Document document = wordprocessingDocument.MainDocumentPart.Document;
            int i = 0;
            bool isInPhrase = false;
            bool isInEndOfPhrase = false;
            foreach (Text text in document.Descendants<Text>()) // <<< Here
            {
                char[] textChars = text.Text.ToCharArray();
                List<WordMatchedPhrase> curParPhrases = matchedPhrases.FindAll(a => (a.firstCharParOccurance.Equals(i) || a.lastCharParOccurance.Equals(i)));
                StringBuilder outStringBuilder = new StringBuilder();

                for (int c = 0; c < textChars.Length; c++)
                {
                    if (isInEndOfPhrase)
                    {
                        isInPhrase = false;
                        isInEndOfPhrase = false;
                    }

                    foreach (var parPhrase in curParPhrases)
                    {
                        if (c == parPhrase.charStartInFirstPar && i == parPhrase.firstCharParOccurance)
                        {
                            outStringBuilder.Append(replaceFor);
                            isInPhrase = true;
                        }
                        if (c == parPhrase.charEndInLastPar && i == parPhrase.lastCharParOccurance)
                        {
                            isInEndOfPhrase = true;
                        }

                    }
                    if (isInPhrase == false && isInEndOfPhrase == false)
                    {
                        outStringBuilder.Append(textChars[c]);
                    }
                }
                text.Text = outStringBuilder.ToString();
                i = i + 1;
            }

            return wordprocessingDocument;
        }

        private static List<WordMatchedPhrase> FindWordMatchedPhrases(WordprocessingDocument wordprocessingDocument, string replaceWhat)
        {
            char[] replaceWhatChars = replaceWhat.ToCharArray();
            int overlapsRequired = replaceWhatChars.Length;
            int overlapsFound = 0;
            int currentChar = 0;
            int firstCharParOccurance = 0;
            int lastCharParOccurance = 0;
            int startChar = 0;
            int endChar = 0;
            List<WordMatchedPhrase> wordMatchedPhrases = new List<WordMatchedPhrase>();
            //
            Document document = wordprocessingDocument.MainDocumentPart.Document;
            int i = 0;
            foreach (Text text in document.Descendants<Text>()) // <<< Here
            {
                char[] textChars = text.Text.ToCharArray();
                for (int c = 0; c < textChars.Length; c++)
                {
                    char compareToChar = replaceWhatChars[currentChar];
                    if (textChars[c] == compareToChar)
                    {
                        currentChar = currentChar + 1;
                        if (currentChar == 1)
                        {
                            startChar = c;
                            firstCharParOccurance = i;
                        }
                        if (currentChar == overlapsRequired)
                        {
                            endChar = c;
                            lastCharParOccurance = i;
                            WordMatchedPhrase matchedPhrase = new WordMatchedPhrase
                            {
                                firstCharParOccurance = firstCharParOccurance,
                                lastCharParOccurance = lastCharParOccurance,
                                charEndInLastPar = endChar,
                                charStartInFirstPar = startChar
                            };
                            wordMatchedPhrases.Add(matchedPhrase);
                            currentChar = 0;
                        }
                    }
                    else
                    {
                        currentChar = 0;

                    }
                }
                i = i + 1;
            }

            return wordMatchedPhrases;

        }

    }
}