using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1.Models
{
    public class Tree
    {
        public Tree(
            int index,
            string species,
            double height,
            double stemDiameter,
            int healthRate,
            double canopyDiameter,
            int locationRate,
            int speciesRate,
            double TreePrice,
            bool isTreeTserifi = false)
        {
            Index = index;
            Species = species;
            Height = height;
            StemDiameter = stemDiameter;
            HealthRate = healthRate;
            CanopyRate = CalculateCanopyRate(canopyDiameter, isTreeTserifi);
            LocationRate = locationRate;
            SpeciesRate = speciesRate;
            PriceInNis = TreePrice;
            SumOfValues = HealthRate + LocationRate + SpeciesRate;
            RootsAreaRadiusInMeters = (stemDiameter * 12) / 100;
            Clonability = SumOfValues <= 12 || HealthRate <= 2 ? "Low" : "High";
        }

        // TODO: add trserifi

        public int Index { get; }

        public int Quatity => 1;

        public string Species { get; set; }

        public double Height { get; set; }

        public double StemDiameter { get; set; }

        public int HealthRate { get; set; }
        
        public int LocationRate { get; set; }
        
        // To be fetched from the agriculture department
        public int SpeciesRate { get; set; }

        public int CanopyRate { get; set; }

        // 0-20  = HealthRate + LocationRate + SpeciesRate
        public int SumOfValues { get; set; }

        // To be fetched from the agriculture department
        public double PriceInNis { get; set; }

        // 12 * StemDiameter / 100
        public double RootsAreaRadiusInMeters { get; set; }

        // if SumOfValues <= 12 || HealthRate <=2 - Low, otherwise High
        public string Clonability { get; set; }

        public string Comments { get; set; }

        public string Status { get; set; }

        // Ask Yarden about the bounds
        private int CalculateCanopyRate(double canopyDiameter, bool isTreeTserifi)
        {
            if (canopyDiameter > 12)
            {
                return 5;
            }

            if (canopyDiameter >= 8)
            {
                return 4;
            }

            if (CanopyRate >= 4)
            {
                return isTreeTserifi ? 5 : 3;
            }

            if (canopyDiameter >= 2)
            {
                return isTreeTserifi ? 4 : 2;
            }

            return isTreeTserifi ? 3 : 1;
        }
    }
}
