namespace Linx
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text.RegularExpressions;

    using ColoredConsole;

    using Microsoft.Office.Core;
    using Microsoft.Office.Interop.PowerPoint;

    using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

    public class DeckIn : InputBase
    {
        private Regex NotesLinkParser = new Regex(@"\b(?:https?://|www\.)\S+\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public override List<Item> ExtractLinks(object file)
        {
            List<Item> results = null;
            Application app = null;
            Presentation presentation = null;

            try
            {
                app = new Application();
                presentation = app.Presentations.Open(file.ToString(), MsoTriState.msoCTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
                results = ParseSlides(presentation);
            }
            catch (Exception ex)
            {
                ColorConsole.WriteLine(ex.Message.White().OnRed());
            }
            finally
            {
                presentation?.Close();
                presentation?.NAR();
                app?.Quit();
                app?.NAR();
            }

            ColorConsole.WriteLine(Environment.NewLine);
            return results;
        }

        private List<Item> ParseSlides(Presentation presentation)
        {
            var results = new List<Item>();
            var s = 1;
            foreach (Slide slide in presentation.Slides)
            {
                ParseSlide(results, s, slide);
                s++;
            }

            return results;
        }

        private void ParseSlide(List<Item> results, int s, Slide slide)
        {
            ColorConsole.Write(s.ToString().Green());
            var links = slide.Hyperlinks;
            if (links.Count > 0)
            {
                ParseLinks(results, s, links.Cast<Hyperlink>()?.Select(link =>
                {
                    var text = link.Type == MsoHyperlinkType.msoHyperlinkRange ? link.TextToDisplay?.Trim() : string.Empty;
                    var item = new Item(string.Empty, text, link.Address);
                    link.NAR();
                    return item;
                })?.ToList());
            }

            if (slide.HasNotesPage == MsoTriState.msoTrue)
            {
                var notesPages = slide.NotesPage;
                foreach (Shape shape in notesPages.Shapes)
                {
                    if (shape.Type == MsoShapeType.msoPlaceholder && shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderBody)
                    {
                        var notes = shape.TextFrame.TextRange.Text;
                        var notesLinks = new List<Item>();

                        // Credit: https://stackoverflow.com/a/10576770
                        foreach (Match m in NotesLinkParser.Matches(notes))
                        {
                            notesLinks.Add(new Item(string.Empty, m.Value.Split('/').LastOrDefault().Replace("-", " ").Replace("_", " "), m.Value));
                        }

                        if (notesLinks.Count > 0)
                        {
                            ParseLinks(results, s, notesLinks);
                        }
                    }
                }
            }


            slide.NAR();
        }
    }
}
