using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DGridLib
{
    public class Localization
    {
        public virtual string saveButton { get; set; }
    }

    public class EnglishLocalization : Localization
    {
        public override string saveButton { get; set; } = "Save";
    }
    public class RussianLocalization : Localization
    {
        public override string saveButton { get; set; } = "Сохранить";
    }
}
