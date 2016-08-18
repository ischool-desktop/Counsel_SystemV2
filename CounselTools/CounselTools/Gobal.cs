using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace CounselTools
{
    public class Gobal
    {
        public static Dictionary<string, List<DataRow>> _multiple_recordDict = new Dictionary<string, List<DataRow>>();
        public static Dictionary<string, List<DataRow>> _priority_dataDict = new Dictionary<string, List<DataRow>>();
        public static Dictionary<string, List<DataRow>> _relativeDict = new Dictionary<string, List<DataRow>>();
        public static Dictionary<string, List<DataRow>> _semester_dataDict = new Dictionary<string, List<DataRow>>();
        public static Dictionary<string, List<DataRow>> _siblingDict = new Dictionary<string, List<DataRow>>();
        public static Dictionary<string, List<DataRow>> _single_recordDict = new Dictionary<string, List<DataRow>>();
        public static Dictionary<string, List<DataRow>> _yearly_dataDict = new Dictionary<string, List<DataRow>>();

    }
}
