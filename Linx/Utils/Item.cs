namespace Linx
{
    public class Item
    {
        public Item(string index, string text, string link)
        {
            this.Index = index;
            this.Text = text;
            this.Link = link;
        }

        public string Index { get; set; }
        public string Text { get; set; }
        public string Link { get; set; }
    }
}
