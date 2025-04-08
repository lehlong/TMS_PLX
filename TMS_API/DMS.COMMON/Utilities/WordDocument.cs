using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace PROJECT.Service.Extention
{
    public class WordDocumentService
    {
        public WordprocessingDocument ReplaceStringInWordDocumennt(WordprocessingDocument doc, string search, string replace)
        {
            var document = doc.MainDocumentPart.Document;
            var texts = document.Descendants<Text>().ToList();
            var fullText = new StringBuilder();
            var offsets = new int[texts.Count];
            for (int i = 0; i < texts.Count; i++) { offsets[i] = fullText.Length; fullText.Append(texts[i].Text); }
            var matches = Regex.Matches(fullText.ToString(), Regex.Escape(search)).Cast<Match>().Reverse();
            foreach (var match in matches)
            {
                int start = match.Index, end = match.Index + match.Length, i = 0;
                while (i < offsets.Length && offsets[i] + texts[i].Text.Length <= start) i++;
                int startIdx = i;
                while (i < offsets.Length && offsets[i] < end) i++;
                int endIdx = i - 1;
                if (startIdx == endIdx)
                {
                    var text = texts[startIdx];
                    int localStart = start - offsets[startIdx];
                    text.Text = text.Text[..localStart] + replace + text.Text[(localStart + search.Length)..];
                }
                else
                {
                    var first = texts[startIdx];
                    int firstCut = start - offsets[startIdx];
                    first.Text = first.Text[..firstCut];
                    for (int j = startIdx + 1; j < endIdx; j++) texts[j].Text = "";
                    var last = texts[endIdx];
                    int lastCut = end - offsets[endIdx];
                    last.Text = replace + last.Text[lastCut..];
                }
            }
            return doc;
        }

        public List<string> FindTextElement(WordprocessingDocument doc) =>
            Regex.Matches(doc.MainDocumentPart.Document.InnerText, @"##\S+@@")
                 .Cast<Match>()
                 .Select(m => m.Value)
                 .ToList();
    }
}
