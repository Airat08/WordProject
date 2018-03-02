using System;
using Microsoft.Office.Interop.Word;

namespace WordProject
{
    interface Style
    {
        bool Equals(Paragraph paragraph);
    }
}
