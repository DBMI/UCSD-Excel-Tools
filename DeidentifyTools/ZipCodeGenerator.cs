using System;
using System.Linq;
using SimpleZipCode;


namespace DeidentifyTools
{

    public class ZipCodeGenerator
    {
        private static readonly IZipCodeRepository _zipCodes = ZipCodeSource.FromMemory().GetRepository();
        private static readonly Random _random = new Random();

        // The code SAYS stateAbbreviation, but seems to work only with the full state name.
        public static string GenerateBogusZipCodeByState(string stateAbbreviation)
        {
            // Search for all ZIP codes within the specified state
            var stateZipCodes = _zipCodes.Search(x => x.State == stateAbbreviation).ToList();

            if (!stateZipCodes.Any())
            {
                throw new ArgumentException($"No ZIP codes found for state abbreviation: {stateAbbreviation}");
            }

            // Select a random ZIP code from the list
            int randomIndex = _random.Next(0, stateZipCodes.Count);
            return stateZipCodes[randomIndex].PostalCode;
        }
    }
}
