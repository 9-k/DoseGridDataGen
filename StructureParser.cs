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
        public List<ValueTuple<DataRow, int>> matches { get; set; }
    }

    public class StructureParserMethods
    {
        public static TG263ParseResult StructureParser(string structName, DataTable TG263Table, int topN=1, float cutoff=0)
        {
            TG263ParseResult result = new TG263ParseResult();
            structName.Trim();
            structName.TrimStart('z', 'Z', '_');
            string sup = structName.ToUpper();
            if (sup.Contains("EVAL"))
            {
                result.IsEval = true;
            }
            if (sup.Contains("-") || sup.Contains("+"))
            {
                result.IsDerived = true;
            }
            // if it starts with 'p' and it matches better if I remove p,
            if (structName.StartsWith("p") && 
                StructureParser(structName.TrimStart('p'), TG263Table).matches.First().Item2 > StructureParser(structName, TG263Table).matches.First().Item2)
            {
                result.IsPlanning = true;
            }

            int numOpts = Regex.Matches(sup, "OPT").Count;
            bool contOptic = sup.Contains("OPTIC");

            // if there's exactly one instance of OPT and the string contains OPTIC
            if (!(numOpts == 1 && contOptic))
            {
                result.IsOpti = true;
            }
            if (sup.Contains("PRV"))
            {
                result.IsPRV = true;
            }

            if (topN >= 1)
            {
                var dists = new List<ValueTuple<string, int>>();
                List<ValueTuple<DataRow, int>> matches = new List<ValueTuple<DataRow, int>>();

                foreach (DataRow row in TG263Table.Rows)
                {
                    // gotta match to upper so that casing doesn't make an issue
                    string TG263name = row["TG263-Primary Name"].ToString().ToUpper();
                    int dist = Fuzz.WeightedRatio(sup, TG263name);
                    dists.Add((TG263name, dist));
                }

                dists.OrderByDescending(x => x.Item2).Take(topN);

                foreach (ValueTuple<string, int> dist in dists)
                {
                    DataRow match = TG263Table.Select($"Name = '{dist.Item1}'").First();
                    matches.Add((match, dist.Item2));
                }
                result.matches = matches;
            }

            // apply cutoff 
            result.matches.RemoveAll(element => element.Item2 < cutoff);
            return result;
        }
    }
}