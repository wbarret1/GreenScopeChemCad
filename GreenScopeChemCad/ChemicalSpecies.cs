using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GreenScopeChemCad
{
 
    [Serializable]
    static class NISTChemicalList
    {
        static private System.Collections.Generic.List<Species> speciesList;
        static NISTChemicalList()
        {
            speciesList = new System.Collections.Generic.List<Species>();
            try {
                System.IO.StringReader reader = new System.IO.StringReader(Properties.Resources.species);
                string nextLine = reader.ReadLine();
                while (nextLine != null)
                {
                    speciesList.Add(new Species(nextLine));
                    nextLine = reader.ReadLine();
                }
            }
            catch (System.Exception obj)
            {
                obj.GetType();
            }
        }

        static public string casNumber(string compoundName)
        {
            string retVal = string.Empty;
            foreach (Species sp in speciesList)
            {
                int result = String.Compare(compoundName, sp.SpeciesName, true);
                if (result == 0)
                {
                    return sp.CASNumber;
                }
            }
            return retVal;
        }
        static public string molecularFormula(string compoundName)
        {
            string retVal = string.Empty;
            foreach (Species sp in speciesList)
            {
                int result = String.Compare(compoundName, sp.SpeciesName, true);
                if (result == 0)
                {
                    return sp.SpeciesFormula;
                }
            }
            return retVal;
        }
    }

    class Species
    {


        public Species(string line)
        {
            char tab = '\t';
            string[]  parts = line.Split(tab);
            SpeciesName = parts[0];
            SpeciesFormula = parts[1];
            CASNumber = parts[2];
        }

        private string m_CASNumber;
        public string CASNumber
        {
            get { return m_CASNumber; }
            set { m_CASNumber = value; }
        }
        private string m_SpeciesName;

        public string SpeciesName
        {
            get { return m_SpeciesName; }
            set { m_SpeciesName = value; }
        }
        private string m_SpeciesFormula;

        public string SpeciesFormula
        {
            get { return m_SpeciesFormula; }
            set { m_SpeciesFormula = value; }
        }
    }
}
