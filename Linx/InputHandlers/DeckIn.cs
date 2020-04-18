namespace Linx
{
    using System;
    using System.Collections.Generic;

    using ColoredConsole;

    using Microsoft.Office.Core;
    using Microsoft.Office.Interop.PowerPoint;

    public class DeckIn : InputBase
    {
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
            ColorConsole.Write(s.ToString().White());
            var links = slide.Hyperlinks;
            if (links.Count > 0)
            {
                ParseLinks(results, s, links);
            }

            slide.NAR();
        }

        private void ParseLinks(List<Item> results, int s, Hyperlinks links)
        {
            var l = 1;
            foreach (Hyperlink link in links)
            {
                var text = string.Empty;
                var address = link.Address?.Sanitize();
                try
                {
                    text = link.TextToDisplay?.Trim();
                }
                catch
                {
                    ColorConsole.Write(l.ToString().OnRed());
                }

                ParseLink(results, s, l, text, address);
                link.NAR();
                l++;
            }

            results.Add(new Item(string.Empty, string.Empty, string.Empty));
            links.NAR();
        }
    }
}
