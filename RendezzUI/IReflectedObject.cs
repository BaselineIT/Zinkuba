using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Rendezz.UI
{
    public interface IReflectedObject<out TImageType>
    {
        TImageType MirrorSource { get; }
    }
}
