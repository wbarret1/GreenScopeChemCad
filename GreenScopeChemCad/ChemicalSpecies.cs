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
            foreach (Species sp in speciesList)
            {
                int result = String.Compare(compoundName, sp.SpeciesName, true);
                if (result == 0)
                {
                    return sp.CASNumber;
                }
            }
            return string.Empty;
        }

        static public string[] NameAndCasNmber(string compoundName)
        {
            string[] retVal = new string[2];
            gov.nih.nlm.chemspell.SpellAidService service = new gov.nih.nlm.chemspell.SpellAidService();
            string response = service.getSugList(compoundName, "All databases");
            var XMLReader = new System.Xml.XmlTextReader(new System.IO.StringReader(response));
            System.Xml.Serialization.XmlSerializer serializer = new System.Xml.Serialization.XmlSerializer(typeof(Synonym));
            if (serializer.CanDeserialize(XMLReader))
            {
                Synonym synonym = (Synonym)serializer.Deserialize(XMLReader);
                foreach (SynonymChemical chemical in synonym.Chemical)
                {
                    int result = String.Compare(compoundName, chemical.Name, true);
                    if (result == 0)
                    {
                        retVal[0] = chemical.CAS;
                        retVal[1] = chemical.Name;
                        return retVal;
                    }
                }
            }
            serializer = new System.Xml.Serialization.XmlSerializer(typeof(SpellAid));
            if (serializer.CanDeserialize(XMLReader))
            {
                SpellAid aid = (SpellAid)serializer.Deserialize(XMLReader);
                bool different = true;
                retVal[0] = aid.Chemical[0].CAS;
                retVal[1] = aid.Chemical[0].Name;
                for (int i = 0; i < aid.Chemical.Length - 1; i++)
                {
                    if (retVal[0] != aid.Chemical[i + 1].CAS)
                    {
                        different = false;
                        retVal[0] = aid.Chemical[i].CAS;
                        retVal[1] = aid.Chemical[i].Name;
                    }
                }
                if (!different)
                {
                    foreach (SpellAidChemical chemical in aid.Chemical)
                    {
                        int result = String.Compare(compoundName, 0, chemical.Name, 0, compoundName.Length, true);
                        if (result == 0 && compoundName.Length >= chemical.Name.Length)
                        {
                            retVal[0] = chemical.CAS;
                            retVal[1] = chemical.Name;
                            return retVal;
                        }
                    }
                SelectChemicalForm form = new SelectChemicalForm(aid, compoundName);
                form.ShowDialog();
                retVal[0] = form.SelectedChemical.CAS;
                retVal[1] = form.SelectedChemical.Name;
                return retVal;
                }
            }
            return retVal;
        }

        static public string molecularFormula(string casNo)
        {
            string retVal = string.Empty;
            foreach (Species sp in speciesList)
            {
                int result = String.Compare(casNo, sp.CASNumber, true);
                if (result == 0)
                {
                    return sp.SpeciesFormula;
                }
            }
            return string.Empty;
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
