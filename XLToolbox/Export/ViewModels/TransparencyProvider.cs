using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Bovender.Mvvm;
using XLToolbox.Export.Models;

namespace XLToolbox.Export.ViewModels
{
    public class TransparencyProvider : EnumProvider<Transparency>
    {
        protected override string GetDescription(Transparency member)
        {
            switch (member)
            {
                case Transparency.TransparentCanvas: return Strings.TransparentCanvas;
                case Transparency.TransparentWhite: return Strings.TransparentWhite;
                case Transparency.WhiteCanvas: return Strings.WhiteCanvas;
                default:
                    throw new InvalidOperationException(
                        "No localized description available for " + member.ToString());
            }
        }

        protected override string GetTooltip(Transparency member)
        {
            switch (member)
            {
                case Transparency.TransparentCanvas: return Strings.TransparentCanvasTooltip;
                case Transparency.TransparentWhite: return Strings.TransparentWhiteTooltip;
                case Transparency.WhiteCanvas: return Strings.WhiteCanvasTooltip;
                default:
                    throw new InvalidOperationException(
                        "No localized tooltip available for " + member.ToString());
            }
        }
    }
}
