using System;
using System.IO;
using System.Data;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Reflection;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using ESAPIUtilities;
using FuzzySharp;
using System.Windows;
using System.Text.RegularExpressions;

namespace DoseGridDataGen
{
    public struct TG263ParseResult
    {
        public bool IsEval { get; set; }
        public bool IsOpti { get; set; }
        public bool IsDerived { get; set; }
        public bool IsPlanning { get; set; }
        public bool IsPRV { get; set; }
        public bool IsCouch { get; set; }
        public List<ValueTuple<DataRow, int>> matches { get; set; }
    }

    public class StructureParserMethods
    {
        public static TG263ParseResult StructureParser(string structName, DataTable TG263Table, int topN=1, int cutoff=0)
        {
            TG263ParseResult result = new TG263ParseResult();
            structName.Trim();
            structName.TrimStart('z', 'Z', '_');
            string sup = structName.ToUpper();
            if (sup.Contains("EVAL")) { result.IsEval = true; }
            if (sup.Contains("-") ||
                sup.Contains("MINUS") ||
                sup.Contains("+") ||
                sup.Contains("PLUS")
                ) { result.IsDerived = true; }
            // if it starts with 'p' and it matches better if I remove p,
            //if (structName.StartsWith("p") && 
            //    StructureParser(structName.TrimStart('p'), TG263Table).matches.First().Item2 > StructureParser(structName, TG263Table).matches.First().Item2)
            //{
            //    result.IsPlanning = true;
            //}

            int numOpts = Regex.Matches(sup, "OPT").Count;
            int numOptics = Regex.Matches(sup, "OPTIC").Count;

            // if the number of matches isn't the same between opt and optic, then one or more of those counts has to be
            // a loose OPT or OPTI ( as #optic matches <= #opt matches, as every #optic match also yields an #opt.
            if (numOpts != numOptics) { result.IsOpti = true; }
            if (sup.Contains("PRV")) { result.IsPRV = true; }
            if (sup.Contains("COUCH")) { result.IsCouch = true; }
            if (topN >= 1)
            {
                var tempMatches = new List<ValueTuple<DataRow, int>>();

                foreach (DataRow row in TG263Table.Rows)
                {
                    // gotta match to upper so that casing doesn't make an issue
                    string TG263name = row["TG263-Primary Name"].ToString();
                    int similarity = Fuzz.WeightedRatio(sup, TG263name.ToUpper());
                    if (similarity > cutoff) { tempMatches.Add((row, similarity)); }
                }
                result.matches = tempMatches.OrderByDescending(x => x.Item2).Take(topN).ToList();
            }

            if (sup.StartsWith("P"))
            {
                int scoreToBeat = StructureParser(structName.TrimStart('p', 'P'), TG263Table).matches.First().Item2;
                if (result.matches.First().Item2 < scoreToBeat)
                {
                    result.IsPlanning = true;
                }
            }
            return result;
        }
    }
}