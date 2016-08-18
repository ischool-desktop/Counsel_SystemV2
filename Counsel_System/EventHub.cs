using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Counsel_System2
{
    internal static class EventHub
    {
        public static event EventHandler CounselChanged;

        public static void OnCounselChanged()
        {
            if (CounselChanged != null)
                CounselChanged(null, EventArgs.Empty);
        }
    }
}
