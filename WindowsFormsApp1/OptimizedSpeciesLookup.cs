using System;
using System.Collections.Generic;
using WindowsFormsApp1.Models;

namespace WindowsFormsApp1
{
    /// <summary>
    /// Optimized species lookup using pre-normalized dictionaries for O(1) performance
    /// Eliminates O(n) linear search bottleneck in species matching
    /// </summary>
    public class OptimizedSpeciesLookup
    {
        private readonly Dictionary<string, TreeSpecie> _exactMatch = new Dictionary<string, TreeSpecie>();
        private readonly Dictionary<string, TreeSpecie> _normalizedMatch = new Dictionary<string, TreeSpecie>();

        public OptimizedSpeciesLookup(Dictionary<string, TreeSpecie> originalSpecies)
        {
            if (originalSpecies == null)
                throw new ArgumentNullException(nameof(originalSpecies));

            // Pre-process all species names for O(1) lookup
            foreach (var kvp in originalSpecies)
            {
                var species = kvp.Value;
                var hebrewName = kvp.Key;

                if (string.IsNullOrEmpty(hebrewName) || species == null)
                    continue;

                // Exact match dictionary
                _exactMatch[hebrewName] = species;

                // Normalized match dictionary (multiple variants)
                var normalized = NormalizeSpeciesName(hebrewName);
                if (!string.IsNullOrEmpty(normalized))
                {
                    _normalizedMatch[normalized] = species;

                    // Add common variations
                    AddVariation(hebrewName.Replace(" ", ""), species);
                    AddVariation(hebrewName.Replace("-", ""), species);
                    AddVariation(hebrewName.Replace(" ", "").Replace("-", ""), species);
                    AddVariation(hebrewName.Replace("-", "").Replace(" ", ""), species);
                }
            }
        }

        private void AddVariation(string variation, TreeSpecie species)
        {
            if (!string.IsNullOrEmpty(variation) && !_normalizedMatch.ContainsKey(variation))
            {
                _normalizedMatch[variation] = species;
            }
        }

        /// <summary>
        /// Attempts to find a species match using O(1) dictionary lookups
        /// </summary>
        /// <param name="inputSpecies">The species name to match</param>
        /// <param name="species">The matched species if found</param>
        /// <returns>True if a match was found, false otherwise</returns>
        public bool TryGetSpecies(string inputSpecies, out TreeSpecie species)
        {
            species = null;

            if (string.IsNullOrWhiteSpace(inputSpecies))
                return false;

            // OPTIMIZATION: O(1) exact lookup first
            if (_exactMatch.TryGetValue(inputSpecies, out species))
                return true;

            // OPTIMIZATION: O(1) normalized lookup
            var normalized = NormalizeSpeciesName(inputSpecies);
            if (!string.IsNullOrEmpty(normalized))
            {
                return _normalizedMatch.TryGetValue(normalized, out species);
            }

            return false;
        }

        /// <summary>
        /// Normalizes species name for consistent matching
        /// </summary>
        /// <param name="name">The species name to normalize</param>
        /// <returns>Normalized species name</returns>
        private static string NormalizeSpeciesName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return string.Empty;

            return name.Trim()
                      .Replace("-", "")
                      .Replace(" ", "")
                      .ToLowerInvariant();
        }

        /// <summary>
        /// Gets the total number of species in the lookup
        /// </summary>
        public int Count => _exactMatch.Count;

        /// <summary>
        /// Gets statistics about the lookup tables
        /// </summary>
        public string GetStatistics()
        {
            return $"Exact matches: {_exactMatch.Count}, Normalized variations: {_normalizedMatch.Count}";
        }
    }
}