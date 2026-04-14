using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace DeidentifyTools
{
    internal class Alias
    {
        private string _alias;
        private bool _isPrincipal;

        // We want to distinguish between the first time we see a name (& decide to hide it)
        // and subsequent similar names that we LINK to the original name.
        // property _isPrincipal is true for that first instance and false for the names linked to it.
        internal Alias(string alias, bool isNew = false)
        {
            this._alias = alias;
            _isPrincipal = isNew;
        }

        internal bool IsPrincipal()
        {
            return _isPrincipal;
        }

        internal string Code()
        {
            return _alias;
        }
    }

    internal class HiddenNames
    {
        private Dictionary<string, Alias> namesAndAliases;

        internal HiddenNames()
        {
            namesAndAliases = new Dictionary<string, Alias>();
        }

        internal void AddName(string newName, string linkedName = "")
        {
            if (string.IsNullOrEmpty(linkedName))
            {
                // Then add this new name.
                if (!namesAndAliases.ContainsKey(newName))
                {
                    // This is a NEW name to be hidden.
                    string hashCode = String.Format("{0:X}", newName.GetHashCode());
                    namesAndAliases.Add(newName, new Alias(alias: "<" + hashCode + ">", isNew: true));
                }
            }
            else
            {
                // Can we find the alias for the linked name?
                string aliasOfLinkedName = GetAlias(linkedName);

                // Put this newName into the dictionary but using the alias for the LINKED name.
                // And mark it as NOT a new alias code.
                if (!namesAndAliases.ContainsKey(newName))
                {
                    namesAndAliases[newName] = new Alias(alias: aliasOfLinkedName, isNew: false);
                }
            }
        }

        // Finds PRINCIPAL names like this one.
        internal List<string> FindSimilarNames(string name = "")
        {
            List<string> similarNames = new List<string>();

            if (string.IsNullOrEmpty(name))
            {
                // It's a signal to send ALL the keys.
                similarNames = namesAndAliases.Where(kvp => kvp.Value.IsPrincipal())
                                    .Select(kvp => kvp.Key)
                                    .ToList<string>();
            }
            else
            {
                double editDistanceThreshold = 0.5;
                double fractionOfWordsPresentThreshold = 0.5;

                Fastenshtein.Levenshtein lev = new Fastenshtein.Levenshtein(name);

                foreach (string key in namesAndAliases.Where(kvp => kvp.Value.IsPrincipal())
                                    .Select(kvp => kvp.Key)
                                    .ToList<string>())
                {
                    // Is every word in this new name present in an existing key?
                    if (Utilities.WordsPresent(name, key) >= fractionOfWordsPresentThreshold)
                    {
                        similarNames.Add(key);
                        continue;
                    }

                    // Test using Levenshtein distance.
                    double wordLength = (double)Math.Min(name.Length, key.Length);
                    int levenshteinDistance = lev.DistanceFrom(key);
                    double relativeDistance = levenshteinDistance / wordLength;

                    if (relativeDistance <= editDistanceThreshold)
                    {
                        similarNames.Add(key);
                    }
                }
            }

            similarNames.Sort();
            return similarNames;
        }

        internal string GetAlias(string name)
        {
            // This only adds name to dictionary if it's not already there.
            AddName(name);

            Alias alias = namesAndAliases[name];
            return alias.Code();
        }

        internal bool HasName(string name)
        {
            return namesAndAliases.ContainsKey(name);
        }
    }
}
