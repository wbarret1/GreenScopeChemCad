using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//// using System.Threading.Tasks;

namespace GreenScopeChemCad
{
    class StreamComponent
    {
        int m_StreamID;
        int m_ComponentPosition;
        int m_ComponentID;
        StreamInfo p_StreamInfo;
        Flowsheet flowsheet;
        CompPPData p_PropertyPackage;
        string m_ComponentName;
        string m_CASNumber;
        string m_MolecularFormula;
        double m_MolecularWeight;
        double m_CriticalTemperature;
        double m_CriticalPressure;
        double m_AcentricFactor;
        double m_BoilingPoint;
        double specificGravityAt60F;
        double m_IdealGasHeatOfFormation;
        double m_IdealGasGibbsFreeEnergyOfFormation;
        int numCarbon;
        int numHydrogen;
        int numNitrogen;
        int numChlorine;
        int numSodium;
        int numOxygen;
        int numPhosphorous;
        int numSulfur;
        string ecClass;
        string rPhrase;
        string MAKcarcinogenCategory;
        string MAKcellMutantGroup;
        string MAKppmValue;
        string MAmgm3Value;
        string bpValue;
        string bpUnit;
        string mpValue;
        string mpUnit;
        string densityValue;
        string densityUnit;
        string relativeDensityValue;
        string vaporPressUnit;
        string vaporPressTemp;
        string vaporPress;
        string flashPtValue;
        string flashPtUnit;
        string nfpaHealth;
        string nfpaFire;
        string nfpaReactivity;
        string logKowStr;
        string m_ERPG2;
        string m_ERPG3;
        string m_IDLH;
        bool hazardous;
        bool triList;
        bool triPBTList;
        int m_CID;
        string m_heatOfCombustion;
        string m_heatOfVaporization;




        public StreamComponent(int StreamID, int ComponentPosition, object vbServer)
        {
            m_StreamID = StreamID;
            m_ComponentPosition = ComponentPosition;
            p_StreamInfo = ((VBServerWrapper)vbServer).GetStreamInfo();
            m_ComponentName = p_StreamInfo.GetComponentNameByPos(m_ComponentPosition);
            m_ComponentID = p_StreamInfo.GetComponentIDByPos(m_ComponentPosition);
            m_CASNumber = NISTChemicalList.casNumber(m_ComponentName);
            if (m_CASNumber == string.Empty)
            {
                string[] values = NISTChemicalList.NameAndCasNmber(m_ComponentName);
                m_CASNumber = values[0];
                m_ComponentName = values[1];
            }
            m_MolecularFormula = NISTChemicalList.molecularFormula(m_CASNumber);
            this.GetCompoundData(m_CASNumber);
            flowsheet = ((VBServerWrapper)vbServer).GetFlowsheet();
            p_PropertyPackage = ((VBServerWrapper)vbServer).GetCompPPData();
            m_MolecularWeight = p_PropertyPackage.GetDataInInternalUnit(m_ComponentPosition, 1);
            m_CriticalTemperature = p_PropertyPackage.GetDataInInternalUnit(m_ComponentPosition, 2) * 5 / 9;
            m_CriticalPressure = p_PropertyPackage.GetDataInInternalUnit(m_ComponentPosition, 3) * 6.89476;
            m_AcentricFactor = p_PropertyPackage.GetDataInInternalUnit(m_ComponentPosition, 4);
            m_BoilingPoint = p_PropertyPackage.GetDataInInternalUnit(m_ComponentPosition, 5) * 5 / 9;
            specificGravityAt60F = p_PropertyPackage.GetDataInInternalUnit(m_ComponentPosition, 6);
            m_IdealGasHeatOfFormation = p_PropertyPackage.GetDataInInternalUnit(m_ComponentPosition, 7) * 2.32601 / 1000;
            m_IdealGasGibbsFreeEnergyOfFormation = p_PropertyPackage.GetDataInInternalUnit(m_ComponentPosition, 8) * 2.32601 / 1000;
            p_PropertyPackage.SaveUserCompData(m_ComponentPosition);
            double value = p_PropertyPackage.GetDataInInternalUnit(m_ComponentPosition, 6);
        }

        //~StreamComponent()
        //{
        //    //if (p_StreamInfo != null)
        //    //    System.Runtime.InteropServices.Marshal.ReleaseComObject(p_StreamInfo);
        //}

        public string ComponentName
        {
            get
            {
                return m_ComponentName;
            }
        }

        public string CASNumber
        {
            get
            {
                return m_CASNumber;
            }
        }

        public string NFPAHealth
        {
            get
            {
                return nfpaHealth;
            }
        }

        public string NFPAFlammability
        {
            get
            {
                return nfpaFire;
            }
        }

        public string NFPAReactivity
        {
            get
            {
                return nfpaReactivity;
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

        public string IDLHvalue
        {
            get
            {
                return m_IDLH;
            }
        }

        public string MAK
        {
            get
            {
                if (!String.IsNullOrEmpty(MAmgm3Value)) return MAmgm3Value;
                if (!String.IsNullOrEmpty(MAKcarcinogenCategory)) return "0.001";
                if (!String.IsNullOrEmpty(MAKcellMutantGroup)) return "0.001";
                return string.Empty;
            }
        }

        public bool Hazarous
        {
            get
            {
                return hazardous;
            }
        }

        public bool IsOnTRIList
        {
            get
            {
                return triList;
            }
        }

        public bool IsPBT
        {
            get
            {
                return triPBTList;
            }
        }

        public string EC_Class
        {
            get
            {
                return ecClass;
            }
        }

        public string RPhrase
        {
            get
            {
                return rPhrase;
            }
        }

        public string HeatOfVaporization
        {
            get
            {
                return m_heatOfVaporization;
            }
        }

        public string Density
        {
            get
            {
                if (!string.IsNullOrEmpty(densityValue)) return densityValue;
                return relativeDensityValue;
            }
        }

        public string VaporPressure
        {
            get
            {
                double tempValue = 0;
                try
                {
                    tempValue = Convert.ToDouble(vaporPress);
                }
                catch(Exception p_Ex){
                    return vaporPress;
                }                
                if (vaporPressUnit == "Pa") return (tempValue * 1e-5).ToString();
                if (vaporPressUnit == "kPa") return (tempValue * 1e-2).ToString();
                return tempValue.ToString();
            }
        }


        public string LogKOW
        {
            get
            {
                return logKowStr;
            }
        }

        public int CarbonAtoms
        {
            get
            {
                return numCarbon;
            }
        }

        public int HydrogenAtoms
        {
            get
            {
                return numHydrogen;
            }
        }

        public int NitrogenAtoms
        {
            get
            {
                return numNitrogen;
            }
        }

        public int ChlorineAtoms
        {
            get
            {
                return numChlorine;
            }
        }

        public int SodiumAtoms
        {
            get
            {
                return numSodium;
            }
        }

        public int OxygenAtoms
        {
            get
            {
                return numOxygen;
            }
        }

        public int Phosphoroustoms
        {
            get
            {
                return numPhosphorous;
            }
        }

        public int SulfurAtoms
        {
            get
            {
                return numSulfur;
            }
        }

        public string MolecularFormula
        {
            get
            {
                return m_MolecularFormula;
            }
        }

        public double MolecularWeight
        {
            get
            {
                return m_MolecularWeight;
            }
        }

        public double CriticalTemperature
        {
            get
            {
                return m_CriticalTemperature;
            }
        }

        public double CriticalPressure
        {
            get
            {
                return m_CriticalPressure;
            }
        }

        public double boilingPoint
        {
            get
            {
                return m_BoilingPoint;
            }
        }

        public double AccentricFactor
        {
            get
            {
                return m_AcentricFactor;
            }
        }

        public double IdealGasHeatOfFormation
        {
            get
            {
                return m_IdealGasHeatOfFormation;
            }
        }

        public double IdealGasGibbsFreeEnergyOfFormation
        {
            get
            {
                return m_IdealGasGibbsFreeEnergyOfFormation;
            }
        }

        public string MeltingPoint
        {
            get
            {
                return mpValue;
            }
        }

        public string FlashPoint
        {
            get
            {
                return flashPtValue;
            }
        }

        public string heatOfCombustion
        {
            get
            {
                return m_heatOfCombustion;
            }
        }

        //public double idealGasHeatCapacity
        //{
        //    get
        //    {
        //        return m_idealGasHeatCapacity;
        //    }
        //}

        //public double heatOfVaporization
        //{
        //    get
        //    {
        //        return m_heatOfVaporization;
        //    }
        //}

        //public double liquidDensity
        //{
        //    get
        //    {
        //        return m_liquidDensity;
        //    }
        //}

        //public double heatOfFormation
        //{
        //    get
        //    {
        //        return 0.0;
        //    }
        //}

        //public double entropyOfFormation
        //{
        //    get
        //    {
        //        return 0.0;
        //    }
        //}

        void GetCompoundData(string casNo)
        {
            ecClass = string.Empty;
            rPhrase = string.Empty;
            MAKcarcinogenCategory = string.Empty;
            MAKcellMutantGroup = string.Empty;
            MAKppmValue = string.Empty;
            MAmgm3Value = string.Empty;
            bpValue = string.Empty;
            bpUnit = string.Empty;
            mpValue = string.Empty;
            mpUnit = string.Empty;
            densityValue = string.Empty;
            densityUnit = string.Empty;
            relativeDensityValue = string.Empty;
            vaporPressUnit = string.Empty;
            vaporPressTemp = string.Empty;
            vaporPress = string.Empty;
            flashPtValue = string.Empty;
            flashPtUnit = string.Empty;
            nfpaHealth = string.Empty;
            nfpaFire = string.Empty;
            nfpaReactivity = string.Empty;
            logKowStr = string.Empty;
            m_ERPG2 = string.Empty;
            m_ERPG3 = string.Empty;
            m_IDLH = string.Empty;
            hazardous = false;
            triList = false;
            triPBTList = false;
            numCarbon = 0;
            numHydrogen = 0;
            numNitrogen = 0;
            numChlorine = 0;
            numSodium = 0;
            numOxygen = 0;
            numPhosphorous = 0;
            numSulfur = 0;
            m_heatOfCombustion = string.Empty;
            m_heatOfVaporization = string.Empty;

            m_ERPG2 = AIHA.ERPG2(casNo);
            m_ERPG3 = AIHA.ERPG3(casNo);
            m_IDLH = IDLH.RevisedIDLH(m_ComponentName);
            hazardous = ListOfLists.IsHAzardous(casNo);
            triList = TRIList.IsTRIChemical(casNo);
            triPBTList = TRIList.IsPBTChemical(casNo);


            string url = "http://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/" + m_ComponentName + "/JSON";
            System.Net.HttpWebRequest request = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(url);
            System.Net.WebResponse response = request.GetResponse();
            System.Runtime.Serialization.Json.DataContractJsonSerializer pugSerializer = new System.Runtime.Serialization.Json.DataContractJsonSerializer(typeof(pugRest.Rootobject));
            pugRest.Rootobject pugChem = (pugRest.Rootobject)pugSerializer.ReadObject(response.GetResponseStream());
            m_CID = pugChem.PC_Compounds[0].id.id.cid;

            foreach (int atom in pugChem.PC_Compounds[0].atoms.element)
            {
                if (atom == 6) numCarbon = numCarbon + 1;
                if (atom == 1) numHydrogen = numHydrogen + 1;
                if (atom == 7) numNitrogen = numNitrogen + 1;
                if (atom == 17) numChlorine = numChlorine + 1;
                if (atom == 11) numSodium = numSodium + 1;
                if (atom == 8) numOxygen = numOxygen + 1;
                if (atom == 15) numPhosphorous = numPhosphorous + 1;
                if (atom == 16) numSulfur = numSulfur + 1;
            }

            string icscNumber = string.Empty;

            System.Collections.Generic.List<string> ICSCnumbers = new System.Collections.Generic.List<string>(0);
            System.Collections.Generic.List<string> caNoss = new System.Collections.Generic.List<string>(0);
            try
            {
                System.IO.StringReader strReader = new System.IO.StringReader(Properties.Resources.ICSCnumberByCAS);
                string nextLine = strReader.ReadLine();
                while (!string.IsNullOrEmpty(nextLine))
                {
                    string[] splits = nextLine.Split('*');
                    ICSCnumbers.Add(splits[0]);
                    caNoss.Add(splits[1].Remove(0, 1));
                    if (splits[1].Remove(0, 1) == casNo) icscNumber = splits[0];
                    nextLine = strReader.ReadLine();
                }
            }
            catch (System.Exception obj)
            {
                obj.GetType();
            }

            if (!string.IsNullOrEmpty(icscNumber))
            {
                url = "http://www.ilo.org/dyn/icsc/showcard.display?p_lang=en&p_card_id=" + icscNumber + "&p_version=1";
                request = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(url);
                response = request.GetResponse();
                System.IO.StreamReader reader = new System.IO.StreamReader(response.GetResponseStream());
                string output = reader.ReadToEnd();

                string pattern = "Symbol: (?<1>\\S+)</span>";
                System.Text.RegularExpressions.Match m = System.Text.RegularExpressions.Regex.Match(output, pattern,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled,
                          TimeSpan.FromSeconds(1));
                ecClass = m.Groups[1].Value;


                pattern = "R: (?<1>\\S+)</span>;";
                m = System.Text.RegularExpressions.Regex.Match(output, pattern,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled,
                          TimeSpan.FromSeconds(1));
                rPhrase = m.Groups[1].Value;

                pattern = "MAK: Carcinogen category: (?<1>\\S+);</span> Germ cell mutagen group: (?<2>\\S+);</span>";
                m = System.Text.RegularExpressions.Regex.Match(output, pattern,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled,
                          TimeSpan.FromSeconds(1));
                MAKcarcinogenCategory = m.Groups[1].Value;
                MAKcellMutantGroup = m.Groups[2].Value;

                pattern = "MAK: (?<1>\\S+) ppm</span>, (?<2>\\S+) mg/m³;";
                m = System.Text.RegularExpressions.Regex.Match(output, pattern,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled,
                          TimeSpan.FromSeconds(1));
                MAKppmValue = m.Groups[1].Value;
                MAmgm3Value = m.Groups[2].Value;

                if (m.Groups.Count == 1)
                {
                    pattern = "MAK \\(respirable fraction\\)</span>: (?<1>\\S+) ppm</span>, (?<2>\\S+) mg/m³;</span>";
                    m = System.Text.RegularExpressions.Regex.Match(output, pattern,
                               System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled,
                              TimeSpan.FromSeconds(1));
                    MAKppmValue = m.Groups[1].Value;
                    MAmgm3Value = m.Groups[2].Value;
                }

                pattern = "Boiling point: (?<1>\\S+)</span>°(?<2>\\S+) <br />";
                m = System.Text.RegularExpressions.Regex.Match(output, pattern,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled,
                          TimeSpan.FromSeconds(1));
                bpValue = m.Groups[1].Value;
                bpUnit = m.Groups[2].Value;

                pattern = "Melting point: (?<1>\\S+)</span>°(?<2>\\S+) <br />";
                m = System.Text.RegularExpressions.Regex.Match(output, pattern,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled,
                          TimeSpan.FromSeconds(1));
                mpValue = m.Groups[1].Value;
                mpUnit = m.Groups[2].Value;

                pattern = "Density: (?<1>\\S+)</span> (?<2>\\S+)<br />";
                m = System.Text.RegularExpressions.Regex.Match(output, pattern,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled,
                          TimeSpan.FromSeconds(1));
                densityValue = m.Groups[1].Value;
                densityUnit = m.Groups[2].Value;

                pattern = "Relative density \\(water = 1\\): (?<1>\\S+)</span>";
                m = System.Text.RegularExpressions.Regex.Match(output, pattern,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled,
                          TimeSpan.FromSeconds(1));
                relativeDensityValue = m.Groups[1].Value;

                pattern = "Vapour pressure, (?<1>\\S+) at (?<2>\\S+)</span>°C: (?<3>\\S+)</span> ";
                m = System.Text.RegularExpressions.Regex.Match(output, pattern,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled,
                          TimeSpan.FromSeconds(1));
                vaporPressUnit = m.Groups[1].Value;
                vaporPressTemp = m.Groups[2].Value;
                vaporPress = m.Groups[3].Value;

                pattern = "Flash point: (?<1>\\S+)</span>°(?<2>\\S+) c.c.<br />";
                m = System.Text.RegularExpressions.Regex.Match(output, pattern,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled,
                          TimeSpan.FromSeconds(1));
                flashPtValue = m.Groups[1].Value;
                flashPtUnit = m.Groups[2].Value;

                pattern = "NFPA Code: H(?<1>\\S+); F(?<2>\\S+); R(?<3>\\S+)</span>";
                m = System.Text.RegularExpressions.Regex.Match(output, pattern,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled,
                          TimeSpan.FromSeconds(1));
                nfpaHealth = m.Groups[1].Value;
                nfpaFire = m.Groups[2].Value;
                nfpaReactivity = m.Groups[3].Value;

                pattern = "Octanol/water partition coefficient as log Pow: (?<1>\\S+)</span>";
                m = System.Text.RegularExpressions.Regex.Match(output, pattern,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled,
                          TimeSpan.FromSeconds(1));
                logKowStr = m.Groups[1].Value;
            }
        }
    }
}

