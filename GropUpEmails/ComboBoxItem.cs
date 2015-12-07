namespace GropUpEmails
{
    public class ComboxItem
    {
        public string Text
        {
            get; set;
        }
        public string Value
        {
            get; set;
        }

        public ComboxItem(string _Text, string _Value)
        {
            Text = _Text;
            Value = _Value;
        }

        public override string ToString()
        {
            return Text;
        }
    }
}
