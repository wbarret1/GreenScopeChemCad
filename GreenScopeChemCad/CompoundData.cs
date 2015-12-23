using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GreenScopeChemCad
{
    static class TRIList
    {
        static private System.Collections.Generic.List<triChemical> triChemicalListData;

        static TRIList()
        {
            triChemicalListData = new System.Collections.Generic.List<triChemical>();
            try
            {
                System.IO.StringReader reader = new System.IO.StringReader(Properties.Resources.tri_chemical_list_for_ry15_11_5_2015_1);
                string nextLine = reader.ReadLine();
                while (nextLine != null)
                {
                    triChemicalListData.Add(new triChemical(nextLine));
                    nextLine = reader.ReadLine();
                }
            }
            catch (System.Exception obj)
            {
                obj.GetType();
            }
        }

        static public bool IsTRIChemical(string CASnumber)
        {
            foreach (triChemical data in triChemicalListData)
            {
                if (CASnumber == data.CasNumber)
                    return true;
            }
            return false;
        }
        static public bool IsPBTChemical(string CASnumber)
        {
            foreach (triChemical data in triChemicalListData)
            {
                if (CASnumber == data.CasNumber)
                {
                    if (data.DeMinimis == "*") return true;
                    return false;
                }
            }
            return false;
        }
    }


    class triChemical
    {
        string m_CasNumber;
        string m_Name;
        string m_DeMinimis;
        string m_CategoryDesription;
        string m_CategoryMember;

        public triChemical(string newLine)
        {
            string[] splits = newLine.Split('\t');
            m_CasNumber = splits[0].Replace(" ", string.Empty);
            m_Name = splits[1].Replace(" ", string.Empty);
            m_DeMinimis = splits[2].Replace(" ", string.Empty);
            m_CategoryDesription = splits[3].Replace(" ", string.Empty);
            m_CategoryMember = splits[4].Replace(" ", string.Empty);
        }

        public string CasNumber
        {
            get
            {
                return m_CasNumber;
            }
        }
        public string DeMinimis
        {
            get
            {
                return m_DeMinimis;
            }
        }
    }

    static class ListOfLists
    {
        static private System.Collections.Generic.List<hazardousSubstance> hazardousSubstanceListData;

        static ListOfLists()
        {
            hazardousSubstanceListData = new System.Collections.Generic.List<hazardousSubstance>();
            try
            {
                System.IO.StringReader reader = new System.IO.StringReader(Properties.Resources.list_of_lists);
                string nextLine = reader.ReadLine();
                while (nextLine != null)
                {
                    hazardousSubstanceListData.Add(new hazardousSubstance(nextLine));
                    nextLine = reader.ReadLine();
                }
            }
            catch (System.Exception obj)
            {
                obj.GetType();
            }
        }

        static public bool IsHAzardous(string CASnumber)
        {
            foreach (hazardousSubstance data in hazardousSubstanceListData)
            {
                if (CASnumber == data.CasNumber)
                    return true;
            }
            return false;
        }
    }


    class hazardousSubstance
    {
        string m_Name;
        string m_NameIndex;
        string m_CasNumber;
        string m_CasNumber313;
        string m_Section302;
        string m_Section304;
        string m_CERCLARQ;
        string m_Section313;
        string m_RCRA;
        string m_CERCLA112;

        public hazardousSubstance(string newLine)
        {
            string[] splits = newLine.Split('\t');
            m_Name = splits[0].Replace(" ", string.Empty);
            m_NameIndex = splits[1].Replace(" ", string.Empty);
            m_CasNumber = splits[2].Replace(" ", string.Empty);
            m_CasNumber313 = splits[3].Replace(" ", string.Empty);
            m_Section302 = splits[4].Replace(" ", string.Empty);
            m_Section304 = splits[5].Replace(" ", string.Empty);
            m_CERCLARQ = splits[6].Replace(" ", string.Empty);
            m_Section313 = splits[7].Replace(" ", string.Empty);
            m_RCRA = splits[8].Replace(" ", string.Empty);
            m_CERCLA112 = splits[9].Replace(" ", string.Empty);
        }

        public string CasNumber
        {
            get
            {
                return m_CasNumber313;
            }
        }
    }

    static class IDLH
    {
        static private System.Collections.Generic.List<nioshSubstance> idlhData;

        static IDLH()
        {
            idlhData = new System.Collections.Generic.List<nioshSubstance>();
            try
            {
                System.IO.StringReader reader = new System.IO.StringReader(Properties.Resources.idlh);
                string nextLine = reader.ReadLine();
                while (nextLine != null)
                {
                    idlhData.Add(new nioshSubstance(nextLine));
                    nextLine = reader.ReadLine();
                }
            }
            catch (System.Exception obj)
            {
                obj.GetType();
            }
        }

        static public string OriginalIDLH(string compoundName)
        {
            foreach (nioshSubstance data in idlhData)
                if (compoundName.ToLower() == data.Substance.ToLower()) return data.OriginalIDLH;
            return string.Empty;
        }

        static public string RevisedIDLH(string compoundName)
        {
            foreach (nioshSubstance data in idlhData)
                if (compoundName.ToLower() == data.Substance.ToLower()) return data.RevisedIDLH;
            return string.Empty;
        }
    }


    class nioshSubstance
    {
        string m_Substance;
        string m_OriginalIDLH;
        string m_RevisedIDLH;
        public nioshSubstance(string newLine)
        {
            string[] splits = newLine.Split('\t');
            m_Substance = splits[0].Split('(')[0];
            m_OriginalIDLH = splits[1];
            m_RevisedIDLH = splits[2];
        }

        public string Substance
        {
            get
            {
                return m_Substance;
            }
        }

        public string OriginalIDLH
        {
            get
            {
                return m_OriginalIDLH;
            }
        }

        public string RevisedIDLH
        {
            get
            {
                return m_RevisedIDLH;
            }
        }
    }

    static class AIHA
    {
        static private System.Collections.Generic.List<ERPGData> aihaData;

        static AIHA()
        {
            aihaData = new System.Collections.Generic.List<ERPGData>();
            try
            {
                System.IO.StringReader reader = new System.IO.StringReader(Properties.Resources._2015_ERPG_Levels);
                string nextLine = reader.ReadLine();
                while (nextLine != null)
                {
                    aihaData.Add(new ERPGData(nextLine));
                    nextLine = reader.ReadLine();
                }
            }
            catch (System.Exception obj)
            {
                obj.GetType();
            }
        }

        //string m_CompoundName;
        //string m_CasNo;
        static public string ERPG1(string CASNumber)
        {
            foreach (ERPGData data in aihaData)
                if (CASNumber == data.CASNumber) return data.ERPG1;
            return string.Empty;
        }

        static public string ERPG2(string CASNumber)
        {
            foreach (ERPGData data in aihaData)
                if (CASNumber == data.CASNumber) return data.ERPG2;
            return string.Empty;
        }

        static public string ERPG3(string CASNumber)
        {
            foreach (ERPGData data in aihaData)
                if (CASNumber == data.CASNumber) return data.ERPG3;
            return string.Empty;
        }

        static public string LEL(string CASNumber)
        {
            foreach (ERPGData data in aihaData)
                if (CASNumber == data.CASNumber) return data.LEL;
            return string.Empty;
        }
    }
    class ERPGData
    {
        string m_CompoundName;
        string m_CasNo;
        string m_ERPG1;
        string m_ERPG2;
        string m_ERPG3;
        string m_LEL;

        public ERPGData(string line)
        {
            char tab = '\t';
            string[] splits = line.Split(tab);
            m_CompoundName = splits[0];
            m_CasNo = splits[1];
            m_ERPG1 = splits[2];
            m_ERPG2 = splits[3];
            m_ERPG3 = splits[4];
            m_LEL = splits[5];
        }

        public string CASNumber
        {
            get
            {
                return m_CasNo;
            }
        }

        public string ERPG1
        {
            get
            {
                return m_ERPG1;
            }
        }

        public string ERPG2
        {
            get
            {
                return m_ERPG2;
            }
        }

        public string ERPG3
        {
            get
            {
                return m_ERPG3;
            }
        }

        public string LEL
        {
            get
            {
                return m_LEL;
            }
        }
    }

    namespace pugRest
    {

        public class Rootobject
        {
            public PC_Compounds[] PC_Compounds { get; set; }
        }

        public class PC_Compounds
        {
            public Id id { get; set; }
            public Atoms atoms { get; set; }
            public Bonds bonds { get; set; }
            public Coord[] coords { get; set; }
            public int charge { get; set; }
            public Prop[] props { get; set; }
            public Count count { get; set; }
        }

        public class Id
        {
            public Id1 id { get; set; }
        }

        public class Id1
        {
            public int cid { get; set; }
        }

        public class Atoms
        {
            public int[] aid { get; set; }
            public int[] element { get; set; }
        }

        public class Bonds
        {
            public int[] aid1 { get; set; }
            public int[] aid2 { get; set; }
            public int[] order { get; set; }
        }

        public class Count
        {
            public int heavy_atom { get; set; }
            public int atom_chiral { get; set; }
            public int atom_chiral_def { get; set; }
            public int atom_chiral_undef { get; set; }
            public int bond_chiral { get; set; }
            public int bond_chiral_def { get; set; }
            public int bond_chiral_undef { get; set; }
            public int isotope_atom { get; set; }
            public int covalent_unit { get; set; }
            public int tautomers { get; set; }
        }

        public class Coord
        {
            public int[] type { get; set; }
            public int[] aid { get; set; }
            public Conformer[] conformers { get; set; }
        }

        public class Conformer
        {
            public float[] x { get; set; }
            public float[] y { get; set; }
            public Style style { get; set; }
        }

        public class Style
        {
            public int[] annotation { get; set; }
            public int[] aid1 { get; set; }
            public int[] aid2 { get; set; }
        }

        public class Prop
        {
            public Urn urn { get; set; }
            public Value value { get; set; }
        }

        public class Urn
        {
            public string label { get; set; }
            public string name { get; set; }
            public int datatype { get; set; }
            public string release { get; set; }
            public string implementation { get; set; }
            public string version { get; set; }
            public string software { get; set; }
            public string source { get; set; }
            public string parameters { get; set; }
        }

        public class Value
        {
            public int ival { get; set; }
            public float fval { get; set; }
            public string binary { get; set; }
            public string sval { get; set; }
        }
    }
}
