using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GropUpEmails
{
    public class ComboBoxItem
    {
        public string Text
        {
            get; set;
        }
        public string Value
        {
            get; set;
        }

        public ComboBoxItem(string _Text, string _Value)
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
