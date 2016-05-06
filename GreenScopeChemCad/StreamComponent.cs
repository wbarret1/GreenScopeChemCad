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
        int m_CID = 0;
        string ecClass = string.Empty;
        string ecClassReference = string.Empty;
        string rPhrase = string.Empty;
        string rPhraseReference = string.Empty;
        string MAKcarcinogenCategory = string.Empty;
        string MAKcellMutantGroup = string.Empty;
        string MAKppmValue = string.Empty;
        string MAmgm3Value = string.Empty;
        string MAKReference = string.Empty;
        string bpValue = string.Empty;
        string bpUnit = string.Empty;
        string bpPressure = string.Empty;
        string bpReference = string.Empty;
        string mpValue = string.Empty;
        string mpUnit = string.Empty;
        string mpReference = string.Empty;
        string densityValue = string.Empty;
        string densityUnit = string.Empty;
        string relativeDensityValue = string.Empty;
        string densityTemperature = string.Empty;
        string densityTemperatureUnit = string.Empty;
        string densityReference = string.Empty;
        string vaporPressUnit = string.Empty;
        string vaporPressTemp = string.Empty;
        string vaporPressTempUnit = string.Empty;
        string vaporPress = string.Empty;
        string vaporPressReference = string.Empty;
        string flashPtValue = string.Empty;
        string flashPtUnit = string.Empty;
        string flashPtReference = string.Empty;
        string logKowStr = string.Empty;
        string nfpaHealth = string.Empty;
        string nfpaHealthReference = string.Empty;
        string nfpaFire = string.Empty;
        string nfpaFireReference = string.Empty;
        string nfpaReactivity = string.Empty;
        string nfpaReactivityReference = string.Empty;
        string logKow = string.Empty;
        string logKowReference = string.Empty;
        string m_ERPG2 = string.Empty;
        string m_ERPG3 = string.Empty;
        string m_IDLH = string.Empty;
        string m_pH = string.Empty;
        string m_pHAdditional = string.Empty;
        string m_pHReference = string.Empty;
        bool hazardous = false;
        bool triList = false;
        bool triPBTList = false;
        int numCarbon = 0;
        int numHydrogen = 0;
        int numNitrogen = 0;
        int numChlorine = 0;
        int numSodium = 0;
        int numOxygen = 0;
        int numPhosphorous = 0;
        int numSulfur = 0;
        string atomsReference = string.Empty;
        string m_HeatOfCombustion = string.Empty;
        string m_HeatOfCombustionUnit = string.Empty;
        string m_HeatOfCombustionConditions = string.Empty;
        string m_HeatOfCombustionReference = string.Empty;
        string m_HeatOfVaporization = string.Empty;
        string m_HeatOfVaporizationUnit = string.Empty;
        string m_HeatOfVaporizationConditions = string.Empty;
        string m_HeatOfVaporizationReference = string.Empty;
        string m_hsdbDocumentURL = string.Empty;
        string m_iloChemicalSafetyCardURL = string.Empty;
        string m_NioshChemicalSafetyCardURL = string.Empty;
        string m_ld50DermalSpecies = string.Empty;
        string m_ld50DermalValue = string.Empty;
        string m_ld50DermalReference = string.Empty;
        string m_ld50OralSpecies = string.Empty;
        string m_ld50OralValue = string.Empty;
        string m_ld50OralReference = string.Empty;
        string m_lc50Value = string.Empty;
        string m_lc50Reference = string.Empty;

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
            flowsheet = ((VBServerWrapper)vbServer).GetFlowsheet();
            p_PropertyPackage = ((VBServerWrapper)vbServer).GetCompPPData();
            m_MolecularWeight = p_PropertyPackage.GetDataInInternalUnit(m_ComponentPosition, 1);
            this.GetCompoundData(m_ComponentName, m_CASNumber);
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
                return m_HeatOfVaporization;
            }
        }

        public string pH
        {
            get
            {
                return m_pH;
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
                catch (Exception p_Ex)
                {
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
                return logKow;
            }
        }
        public string LD50DermalSpecies
        {
            get
            {
                return m_ld50DermalSpecies;
            }
        }

        public string LD50DermalValue
        {
            get
            {
                return m_ld50DermalValue;
            }
        }

        public string LD50DermalReference
        {
            get
            {
                return m_ld50DermalReference;
            }
        }

        public string LD50OralSpecies
        {
            get
            {
                return m_ld50OralSpecies;
            }
        }

        public string LD50OralValue
        {
            get
            {
                return m_ld50OralValue;
            }
        }

        public string LD50OralReference
        {
            get
            {
                return m_ld50OralReference;
            }
        }

        public string LC50Value
        {
            get
            {
                return m_lc50Value;
            }
        }

        public string LC50Reference
        {
            get
            {
                return m_lc50Reference;
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

         public string boilingPoint
        {
            get
            {
                if (bpUnit == "C") return bpValue;
                return bpUnit;
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
                if (mpUnit == "C") return mpValue;
                return mpUnit;
            }
        }

        public string FlashPoint
        {
            get
            {
                if (flashPtUnit == "C") return flashPtValue;
                return flashPtUnit;
            }
        }

        public string heatOfCombustionKgPerKg
        {
            get
            {
                if (m_HeatOfCombustionUnit.ToLower() == "kj/g")
                {
                    return (Convert.ToDouble(m_HeatOfCombustion) * 1000).ToString();
                }
                if (m_HeatOfCombustionUnit.ToLower() == "btu/lb")
                {
                    return (Convert.ToDouble(m_HeatOfCombustion) * 2.32599999962).ToString();
                }
                if (m_HeatOfCombustionUnit.ToLower() == "j/kmol")
                {
                    return (Convert.ToDouble(m_HeatOfCombustion) * 1000 * 1000 * m_MolecularWeight).ToString();
                }
                if (m_HeatOfCombustionUnit.ToLower() == "kj/mole" || m_HeatOfCombustionUnit.ToLower() == "kj/mol")
                {
                    return (Convert.ToDouble(m_HeatOfCombustion) * 1000 * m_MolecularWeight).ToString();
                }
                if (m_HeatOfCombustionUnit.ToLower() == "kj/kg")
                {
                    return m_HeatOfCombustion;
                }
                if (string.IsNullOrEmpty(m_HeatOfCombustion) && string.IsNullOrEmpty(m_HeatOfCombustionUnit))
                {
                    return string.Empty; ;
                }
                throw new System.Exception("The unit " + m_HeatOfCombustionUnit + " is not currently available.");
            }
        }

        public string heatOfCombustion
        {
            get
            {
                return m_HeatOfCombustion;
            }
        }

        //public double idealGasHeatCapacity
        //{
        //    get
        //    {
        //        return m_idealGasHeatCapacity;
        //    }
        //}
        public string heatOfVaporizationKjPerKg
        {
            get
            {
                if (m_HeatOfVaporizationUnit.ToLower() == "kj/kg")
                {
                    return m_HeatOfVaporization;
                }
                if (m_HeatOfVaporizationUnit.ToLower() == "kj/g")
                {
                    return (Convert.ToDouble(m_HeatOfVaporization) * 1000).ToString();
                }
                if (m_HeatOfVaporizationUnit.ToLower() == "btu/lb")
                {
                    return (Convert.ToDouble(m_HeatOfVaporization) * 2.32599999962).ToString();
                }
                if (m_HeatOfVaporizationUnit.ToLower() == "g-cal/g")
                {
                    return (Convert.ToDouble(m_HeatOfVaporization) * 4.184 * 1000).ToString();
                }
                if (m_HeatOfVaporizationUnit.ToLower() == "kcal/mole")
                {
                    return (Convert.ToDouble(m_HeatOfVaporization) * 4.184 * m_MolecularWeight).ToString();
                }
                if (m_HeatOfVaporizationUnit.ToLower() == "kj/mole" || m_HeatOfVaporizationUnit.ToLower() == "kj/mol")
                {
                    return (Convert.ToDouble(m_HeatOfVaporization) * 1000 * m_MolecularWeight).ToString();
                }
                if (m_HeatOfVaporizationUnit.ToLower() == "gcal/gmole")
                {
                    return (Convert.ToDouble(m_HeatOfVaporization) * 4.184 * m_MolecularWeight * 1000).ToString();
                }
                if (m_HeatOfVaporizationUnit.ToLower() == "cal/g")
                {
                    return (Convert.ToDouble(m_HeatOfVaporization) * 4.184).ToString();
                }
                if (string.IsNullOrEmpty(m_HeatOfVaporization) && string.IsNullOrEmpty(m_HeatOfVaporizationUnit))
                {
                    return string.Empty; ;
                }
                throw new System.Exception("The heat of vaporization unit " + m_HeatOfVaporizationUnit + " is not currently available.");
            }
        }

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

        void GetCompoundData(string compoundName, string casNo)
        {
            m_ERPG2 = AIHA.ERPG2(casNo);
            m_ERPG3 = AIHA.ERPG3(casNo);
            m_IDLH = IDLH.RevisedIDLH(compoundName);
            hazardous = ListOfLists.IsHAzardous(casNo);
            triList = TRIList.IsTRIChemical(casNo);
            triPBTList = TRIList.IsPBTChemical(casNo);


            atomsReference = "http://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/" + compoundName + "/JSON";
            System.Net.HttpWebRequest request = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(atomsReference);
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
                    string[] split = nextLine.Split('*');
                    ICSCnumbers.Add(split[0]);
                    caNoss.Add(split[1].Remove(0, 1));
                    if (split[1].Remove(0, 1) == casNo) icscNumber = split[0];
                    nextLine = strReader.ReadLine();
                }
            }
            catch (System.Exception obj)
            {
                obj.GetType();
            }

            if (!string.IsNullOrEmpty(icscNumber))
            {
                m_iloChemicalSafetyCardURL = "http://www.ilo.org/dyn/icsc/showcard.display?p_lang=en&p_card_id=" + icscNumber + "&p_version=1";
                m_NioshChemicalSafetyCardURL = "http://www.cdc.gov/niosh/ipcsneng/neng" + icscNumber + ".html";

                request = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(m_iloChemicalSafetyCardURL);
                response = request.GetResponse();
                System.IO.StreamReader reader = new System.IO.StreamReader(response.GetResponseStream());
                string output = reader.ReadToEnd();

                string pattern1 = "Symbol: (?<1>\\S+)</span>";
                System.Text.RegularExpressions.Match m1 = System.Text.RegularExpressions.Regex.Match(output, pattern1,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled
                           );
                ecClass = m1.Groups[1].Value;
                ecClassReference = m_iloChemicalSafetyCardURL;


                pattern1 = "R: (?<1>\\S+)</span>;";
                m1 = System.Text.RegularExpressions.Regex.Match(output, pattern1,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
                rPhrase = m1.Groups[1].Value;
                rPhraseReference = m_iloChemicalSafetyCardURL;

                pattern1 = "MAK: Carcinogen category: (?<1>\\S+);</span> Germ cell mutagen group: (?<2>\\S+);</span>";
                m1 = System.Text.RegularExpressions.Regex.Match(output, pattern1,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
                MAKcarcinogenCategory = m1.Groups[1].Value;
                MAKcellMutantGroup = m1.Groups[2].Value;

                pattern1 = "MAK: (?<1>\\S+) ppm</span>, (?<2>\\S+) mg/m³;";
                m1 = System.Text.RegularExpressions.Regex.Match(output, pattern1,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
                MAKppmValue = m1.Groups[1].Value;
                MAmgm3Value = m1.Groups[2].Value;

                if (m1.Groups.Count == 1)
                {
                    pattern1 = "MAK \\(respirable fraction\\)</span>: (?<1>\\S+) ppm</span>, (?<2>\\S+) mg/m³;</span>";
                    m1 = System.Text.RegularExpressions.Regex.Match(output, pattern1,
                               System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
                    MAKppmValue = m1.Groups[1].Value;
                    MAmgm3Value = m1.Groups[2].Value;
                }
                MAKReference = m_iloChemicalSafetyCardURL;

                pattern1 = "Boiling point: (?<1>\\S+)</span>°(?<2>\\S+) <br />";
                m1 = System.Text.RegularExpressions.Regex.Match(output, pattern1,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
                bpValue = m1.Groups[1].Value;
                bpUnit = m1.Groups[2].Value;
                bpReference = m_iloChemicalSafetyCardURL;

                pattern1 = "Melting point: (?<1>\\S+)</span>°(?<2>\\S+) <br />";
                m1 = System.Text.RegularExpressions.Regex.Match(output, pattern1,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
                mpValue = m1.Groups[1].Value;
                mpUnit = m1.Groups[2].Value;
                mpReference = m_iloChemicalSafetyCardURL;

                pattern1 = "Density: (?<1>\\S+)</span> (?<2>\\S+)<br />";
                m1 = System.Text.RegularExpressions.Regex.Match(output, pattern1,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
                densityValue = m1.Groups[1].Value;
                densityUnit = m1.Groups[2].Value;

                pattern1 = "Relative density \\(water = 1\\): (?<1>\\S+)</span>";
                m1 = System.Text.RegularExpressions.Regex.Match(output, pattern1,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
                relativeDensityValue = m1.Groups[1].Value;
                densityReference = m_iloChemicalSafetyCardURL;

                pattern1 = "Vapour pressure, (?<1>\\S+) at (?<2>\\S+)</span>°C: (?<3>\\S+)</span> ";
                m1 = System.Text.RegularExpressions.Regex.Match(output, pattern1,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
                vaporPressUnit = m1.Groups[1].Value;
                vaporPressTemp = m1.Groups[2].Value;
                vaporPress = m1.Groups[3].Value;
                vaporPressReference = m_iloChemicalSafetyCardURL;

                pattern1 = "Flash point: (?<1>\\S+)</span>°(?<2>\\S+) c.c.<br />";
                m1 = System.Text.RegularExpressions.Regex.Match(output, pattern1,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
                flashPtValue = m1.Groups[1].Value;
                flashPtUnit = m1.Groups[2].Value;
                flashPtReference = m_iloChemicalSafetyCardURL;

                pattern1 = "NFPA Code: H(?<1>\\S+); F(?<2>\\S+); R(?<3>\\S+)</span>";
                m1 = System.Text.RegularExpressions.Regex.Match(output, pattern1,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
                nfpaHealth = m1.Groups[1].Value;
                nfpaHealthReference = m_iloChemicalSafetyCardURL;
                nfpaFire = m1.Groups[2].Value;
                nfpaFireReference = m_iloChemicalSafetyCardURL;
                nfpaReactivity = m1.Groups[3].Value;
                nfpaReactivityReference = m_iloChemicalSafetyCardURL;

                pattern1 = "Octanol/water partition coefficient as log Pow: (?<1>\\S+)</span>";
                m1 = System.Text.RegularExpressions.Regex.Match(output, pattern1,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
                logKow = m1.Groups[1].Value;
                logKowReference = m_iloChemicalSafetyCardURL;
            }

            // http://toxnet.nlm.nih.gov/cgi-bin/sis/search2/f?./temp/~oiB60G:1
            string uriString = "http://toxnet.nlm.nih.gov/cgi-bin/sis/search2";
            request = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(uriString);
            string postData = "queryxxx=" + casNo;
            postData += "&chemsyn=1";
            postData += "&database=hsdb";
            postData += "&Stemming=1";
            postData += "&and=1";
            postData += "&second_search=1";
            postData += "&gateway=1";
            var data = Encoding.ASCII.GetBytes(postData);
            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            request.ContentLength = data.Length;
            using (var stream = request.GetRequestStream())
            {
                stream.Write(data, 0, data.Length);
            }
            response = (System.Net.HttpWebResponse)request.GetResponse();
            string responseString = new System.IO.StreamReader(response.GetResponseStream()).ReadToEnd();
            string s1 = responseString.Replace("<br>", "");
            System.Xml.XmlDocument document = new System.Xml.XmlDocument();
            document.Load(new System.IO.StringReader(s1));
            string tempFileName = document.FirstChild["TemporaryFile"].InnerText;
            uriString = "http://toxnet.nlm.nih.gov/cgi-bin/sis/search2/f?" + tempFileName;

            //// Whole Document
            m_hsdbDocumentURL = "http://toxgate.nlm.nih.gov/cgi-bin/sis/search2/r?dbs+hsdb:@term+@DOCNO+" + document.FirstChild["Id"].InnerText.Split(' ')[0];
            //request = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(m_hsdbDocumentURL);
            //response = (System.Net.HttpWebResponse)request.GetResponse();
            //System.IO.StringReader full = new System.IO.StringReader(new System.IO.StreamReader(response.GetResponseStream()).ReadToEnd());

            // Chemical Properties 
            request = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(uriString + ":1:cpp");
            response = (System.Net.HttpWebResponse)request.GetResponse();
            string propertiesResposne = new System.IO.StreamReader(response.GetResponseStream()).ReadToEnd();

            // pH
            string pattern = "pH:\\s*(?<1>\\S+)\\s*(?<2>[^\\n]*)\\s*<NOINDEX>(?<3>[^\\n]*)</NOINDEX>";
            System.Text.RegularExpressions.Match m = System.Text.RegularExpressions.Regex.Match(propertiesResposne, pattern,
                       System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
            m_pH = m.Groups[1].Value;
            m_pHAdditional = m.Groups[2].Value;
            m_pHReference = m.Groups[3].Value;

            // heat of vaporization: 
            pattern = "Heat of Vaporization:</h3>\\s*<br>\\s*Latent:\\s*(?<1>\\S+)\\s*(?<2>\\S+)\\s*(?<4>[^\\n]*)<br><code><NOINDEX>(?<3>[^\\n]*)</NOINDEX>"; ;
            m = System.Text.RegularExpressions.Regex.Match(propertiesResposne, pattern,
                       System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
            if (m.Groups.Count == 1)
            {
                pattern = "Heat of Vaporization:</h3>\\s*<br>\\s*Enthalpy of vaporization:\\s*(?<1>\\S+)\\s*(?<2>\\S+)\\s*(?<4>[^\\n]*)<br><code><NOINDEX>(?<3>[^\\n]*)</NOINDEX>";
                m = System.Text.RegularExpressions.Regex.Match(propertiesResposne, pattern,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
            }
            if (m.Groups.Count == 1)
            {
                pattern = "Heat of Vaporization:</h3>\\s*<br>\\s*(?<1>\\S+)\\s*(?<2>\\S+)\\s*(?<4>[^\\n]*)<br><code><NOINDEX>(?<3>[^\\n]*)</NOINDEX>";
                m = System.Text.RegularExpressions.Regex.Match(propertiesResposne, pattern,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
            }
            m_HeatOfVaporization = m.Groups[1].Value;
            m_HeatOfVaporizationUnit = m.Groups[2].Value;
            m_HeatOfVaporizationReference = m.Groups[3].Value;
            if (m_CASNumber == "64-19-7")
            {
                m_HeatOfVaporization = "1219055.6";
                m_HeatOfVaporizationUnit = "kJ/kg";
                m_HeatOfVaporizationReference = "[Haynes, W.M. (ed.). CRC Handbook of Chemistry and Physics. 94th Edition. CRC Press LLC, Boca Raton: FL 2013-2014, p. 6-132] **PEER REVIEWED** ";
            }

            // heat of combustion: 
            pattern = "Heat of Combustion:</h3>\\s*<br>\\s*(?<1>\\S+)\\s*(?<2>\\S+)\\s*<br><code><NOINDEX>(?<3>[^\\n]*)</NOINDEX>";
            m = System.Text.RegularExpressions.Regex.Match(propertiesResposne, pattern,
                       System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
            if (m.Groups.Count == 1)
            {
                pattern = "Heat of Combustion:</h3>\\s*<br>\\s*(?<1>\\S+)\\s*(?<2>\\S+)\\s*(?<3>[^\\n]*)\\s*<br><code><NOINDEX>(?<3>[^\\n]*)</NOINDEX>";
                m = System.Text.RegularExpressions.Regex.Match(propertiesResposne, pattern,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
            }
            m_HeatOfCombustion = m.Groups[1].Value;
            m_HeatOfCombustionUnit = m.Groups[2].Value;
            m_HeatOfCombustionReference = m.Groups[3].Value;

            // Octanol Water Partitioning Coefficient
            if (string.IsNullOrEmpty(logKow))
            {
                // owpc: 
                pattern = "log\\s*Kow\\s*=\\s*(?<1>\\S+)<br><code><NOINDEX>(?<2>[^\\n]*)</NOINDEX>";
                m = System.Text.RegularExpressions.Regex.Match(propertiesResposne, pattern,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
                logKow = m.Groups[1].Value;
                logKowReference = m.Groups[2].Value;
            }

            // Flash Point: 
            if (string.IsNullOrEmpty(flashPtValue))
            {
                // Chemical Properties 
                pattern = "Flash Point:</h3>\\s*<br>\\s*(?<1>\\S+)\\s*deg\\s*(?<2>\\S+),\\s*(?<5>\\S+)\\s*deg\\s*(?<6>\\S+)\\s*(?<4>[^\\n]*)\\s*<br><code><NOINDEX>(?<3>[^\\n]*)</NOINDEX>";
                m = System.Text.RegularExpressions.Regex.Match(propertiesResposne, pattern,
                               System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
                if (m.Groups.Count == 1)
                {
                    pattern = "Flash Point:</h3>\\s*<br>\\s*(?<1>\\S+)\\s*deg\\s*(?<2>\\S+)\\s*(?<4>[^\\n]*)\\s*<br><code><NOINDEX>(?<3>[^\\n]*)</NOINDEX>";
                    m = System.Text.RegularExpressions.Regex.Match(propertiesResposne, pattern,
                                   System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
                }
                if (m.Groups[2].Value == "C")
                {
                    flashPtValue = m.Groups[1].Value;
                    flashPtUnit = m.Groups[2].Value;
                }
                else if (m.Groups[4].Value == "F")
                {
                    flashPtValue = ((Convert.ToDouble(m.Groups[1].Value) - 32) * 5 / 9).ToString();
                    flashPtUnit = "C";
                }
                flashPtReference = m.Groups[3].Value;
            }

            // Vapor Pressure: 
            if (string.IsNullOrEmpty(vaporPress))
            {
                pattern = "Vapor Pressure:</h3>\\s*<br><\\s*(?<1>\\S+)\\s*(?<2>[^\"']*) at (?<3>\\S+) deg (?<4>\\S+)\\s*(?<5>[^\\n]*)\\s*<br><code><NOINDEX>(?<3>[^\\n]*)</NOINDEX>";
                m = System.Text.RegularExpressions.Regex.Match(propertiesResposne, pattern,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
                if (m.Groups.Count == 1)
                {
                    pattern = "Vapor Pressure:</h3>\\s*<br>\\s*(?<1>\\S+)\\s*(?<2>[^\"']*) at (?<3>\\S+) deg (?<4>\\S+)\\s*(?<5>[^\\n]*)\\s*<br><code><NOINDEX>(?<3>[^\\n]*)</NOINDEX>";
                    m = System.Text.RegularExpressions.Regex.Match(propertiesResposne, pattern,
                               System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
                }
                vaporPress = m.Groups[1].Value.ToLower();
                if (vaporPress.Contains("x"))
                {
                    vaporPress = vaporPress.Replace("x10", "x");
                    string[] splits = vaporPress.Split('x');
                    vaporPress = (Convert.ToDouble(splits[0]) * Math.Pow(10, Convert.ToDouble(splits[1]))).ToString();
                }
                vaporPressUnit = m.Groups[2].Value;
                vaporPressTemp = m.Groups[3].Value;
                vaporPressTempUnit = m.Groups[4].Value;
            }

            // boilingPoint: 
            if (string.IsNullOrEmpty(bpValue))
            {
                pattern = "Boiling Point:</h3>\\s*<br>\\s*(?<1>\\S+) deg (?<2>\\S+)<br><code><NOINDEX>(?<3>[^\\n]*)</NOINDEX>";
                m = System.Text.RegularExpressions.Regex.Match(propertiesResposne, pattern,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
                if (m.Groups.Count == 1)
                {
                    pattern = "Boiling Point:</h3>\\s*<br>\\s*(?<1>\\S+) deg (?<2>\\S+)\\s*(?<4>[^\\n]*)<br><code><NOINDEX>(?<3>[^\\n]*)</NOINDEX>";
                    m = System.Text.RegularExpressions.Regex.Match(propertiesResposne, pattern,
                               System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
                }
                bpValue = m.Groups[1].Value;
                bpUnit = m.Groups[2].Value;
                bpReference = m.Groups[3].Value;
            }

            // meltingPoint: 
            if (string.IsNullOrEmpty(mpValue))
            {
                pattern = "Melting Point:</h3>\\s*<br>\\s*(?<1>\\S+)\\s*deg\\s*(?<2>\\S+)\\s*(?<4>[^\\n]*)\\s*<br><code><NOINDEX>(?<3>[^\\n]*)</NOINDEX>";
                m = System.Text.RegularExpressions.Regex.Match(propertiesResposne, pattern,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
                mpValue = m.Groups[1].Value;
                mpUnit = m.Groups[2].Value;
                mpReference = m.Groups[3].Value;
            }

            // density: 
            if (string.IsNullOrEmpty(densityValue) || string.IsNullOrEmpty(relativeDensityValue))
            {
                pattern = "Density/Specific Gravity:</h3>\\s*<br>\\s*Gas:\\s*(?<1>\\S+)\\s*(?<2>[^\"']*)\\s*at\\s*(?<4>[^\\n]*)\\s*<br><code><NOINDEX>(?<4>[^\\n]*)</NOINDEX>";
                m = System.Text.RegularExpressions.Regex.Match(propertiesResposne, pattern,
                           System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
                if (m.Groups.Count == 1)
                {
                    pattern = "Density/Specific Gravity:</h3>\\s*<br>\\s*Absolute density:\\s*(?<1>\\S+)\\s*(?<2>[^\"']*)\\s*at\\s*(?<4>[^\\n]*)\\s*<br><code><NOINDEX>(?<4>[^\\n]*)</NOINDEX>";
                    m = System.Text.RegularExpressions.Regex.Match(propertiesResposne, pattern,
                               System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
                }
                if (m.Groups.Count == 1)
                {
                    pattern = "Density/Specific Gravity:</h3>\\s*<br>\\s*(?<1>\\S+)\\s*(?<2>[^\"']*)\\s*at\\s*(?<4>[^\\n]*)\\s*<br><code><NOINDEX>(?<4>[^\\n]*)</NOINDEX>";
                    m = System.Text.RegularExpressions.Regex.Match(propertiesResposne, pattern,
                               System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
                }
                densityValue = m.Groups[1].Value;
                densityUnit = m.Groups[2].Value;
                densityTemperature = m.Groups[3].Value;
                densityTemperatureUnit = m.Groups[4].Value;
                densityReference = m.Groups[5].Value;
            }

            // EcoTox
            request = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(uriString + ":1:etxv");
            response = (System.Net.HttpWebResponse)request.GetResponse();
            string ecoToxResposne = new System.IO.StreamReader(response.GetResponseStream()).ReadToEnd();
            pattern = "LC50; Species:\\s*(?<1>[^\\n]*)\\s*[,;] Conditions:\\s*(?<2>[^\\n]*)\\s*[,;] Concentration:\\s*(?<3>\\S+)\\s*(?<4>\\S+)\\s*for\\s*(?<5>[^\\n]*)\\s*<br><code><NOINDEX>\\s*(?<6>[^\\n]*)\\s*</NOINDEX>";
            System.Text.RegularExpressions.MatchCollection matchColl = System.Text.RegularExpressions.Regex.Matches(ecoToxResposne, pattern,
                System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);

            double lc50 = 1000000.0;
            System.Text.RegularExpressions.Match lc50Match = null;
            foreach (System.Text.RegularExpressions.Match match in matchColl)
            {
                string lc50STR = match.Groups[3].Value;
                lc50STR = lc50STR.Replace("&gt;", "");
                lc50STR = lc50STR.Replace("&lt;", "");
                double conc = Convert.ToDouble(lc50STR);
                string unit = match.Groups[4].Value;
                if (string.Compare(unit, "mg/l", true) == 0)
                {
                    if (conc < lc50)
                    {
                        lc50 = conc;
                        lc50Match = match;
                    }
                }
                else if (string.Compare(unit, "ug/l", true) == 0)
                {
                    {
                        if (conc / 1000 < lc50) lc50 = conc / 1000;
                        lc50Match = match;

                    }
                }
            }
            if (lc50Match != null)
            {
                m_lc50Value = lc50Match.Groups[3].Value;
                m_lc50Reference = "Species: " + lc50Match.Groups[1].Value + "; Conditions: " + lc50Match.Groups[2].Value + "; Time: " + lc50Match.Groups[5].Value + "; Reference: " + lc50Match.Groups[6].Value;
            }
            //pattern = "EC50; Species:\\s*(?<1>[^\\n]*)\\s*[,;] Conditions:\\s*(?<2>[^\\n]*)\\s*[,;] Concentration:\\s*(?<3>[^\\n]*)\\s*for\\s*(?<4>[^\\n]*)\\s*[,;] Effect: \\s*(?<5>[^\\n]*)\\s*<br><code><NOINDEX>\\s*(?<6>[^\\n]*)\\s*</NOINDEX>";
            //matchColl = System.Text.RegularExpressions.Regex.Matches(ecoToxResposne, pattern,
            //    System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);

            // NonHuman Tox
            request = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(uriString + ":1:ntxv");
            response = (System.Net.HttpWebResponse)request.GetResponse();
            string nonHumanToxResposne = new System.IO.StreamReader(response.GetResponseStream()).ReadToEnd();

            pattern = "LD50\\s*(?<1>\\S+)\\s*oral\\s*[(&gt;)\\s*]\\s*(?<2>\\S+)\\s*(?<3>\\S+)\\s*<br><code><NOINDEX>\\s*(?<4>[^\\n]*)\\s*</NOINDEX>";
            matchColl = System.Text.RegularExpressions.Regex.Matches(nonHumanToxResposne, pattern,
                System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);

            double ld50Oral = 1000000.0;
            System.Text.RegularExpressions.Match ld50OralMatch = null;
            foreach (System.Text.RegularExpressions.Match match in matchColl)
            {
                string ld50STR = match.Groups[2].Value;
                ld50STR = ld50STR.Replace("&gt;", "");
                ld50STR = ld50STR.Replace("&lt;", "");
                double conc = 0;
                if (!double.TryParse(ld50STR, out conc))
                {
                    string[] values = ld50STR.Split('-');
                    double.TryParse(values[0], out conc);
                }
                string unit = match.Groups[3].Value;
                if (string.Compare(unit, "mg/kg", true) == 0)
                {
                    if (conc < ld50Oral)
                    {
                        ld50Oral = conc;
                        ld50OralMatch = match;
                    }
                }
                else if (string.Compare(unit, "g/kg", true) == 0)
                {
                    {
                        if (conc * 1000 < ld50Oral) ld50Oral = conc * 1000;
                        ld50OralMatch = match;

                    }
                }
            }
            if (ld50OralMatch != null)
            {
                m_ld50OralSpecies = ld50OralMatch.Groups[1].Value;
                m_ld50OralValue = ld50Oral.ToString();
                m_ld50OralReference = ld50OralMatch.Groups[4].Value;
            }

            double ld50Dermal = 1000000.0;
            System.Text.RegularExpressions.Match ld50DermalMatch = null;
            string[] dermalSynomoyms =
            {
                "dermal",
                "percutaneous",
                "sc",
                "skin"
            };
            foreach (string synonym in dermalSynomoyms)
            {
                pattern = "LD50\\s*(?<1>\\S+)\\s*" + synonym + "\\s*[(&gt;)\\s*]\\s*(?<2>\\S+)\\s*(?<3>\\S+)\\s*<br><code><NOINDEX>\\s*(?<4>[^\\n]*)\\s*</NOINDEX>";
                matchColl = System.Text.RegularExpressions.Regex.Matches(nonHumanToxResposne, pattern,
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);

                foreach (System.Text.RegularExpressions.Match match in matchColl)
                {
                    string ld50STR = match.Groups[2].Value;
                    ld50STR = ld50STR.Replace("&gt;", "");
                    ld50STR = ld50STR.Replace("&lt;", "");
                    double conc = Convert.ToDouble(ld50STR);
                    string unit = match.Groups[3].Value;
                    if (string.Compare(unit, "mg/kg", true) == 0)
                    {
                        if (conc < ld50Dermal)
                        {
                            ld50Dermal = conc;
                            ld50DermalMatch = match;
                        }
                    }
                    else if (string.Compare(unit, "g/kg", true) == 0)
                    {
                        {
                            if (conc * 1000 < ld50Dermal) ld50Dermal = conc * 1000;
                            ld50DermalMatch = match;

                        }
                    }
                }
            }
            pattern = "LD50\\s*(?<1>\\w*)\\s*[(&gt;)\\s*]\\s*(?<2>\\S+)\\s*(?<3>\\S+)\\s*dermal[.\\s+]\\s*\\s*<br><code><NOINDEX>\\s*(?<4>[^\\n]*)\\s*</NOINDEX>";
            matchColl = System.Text.RegularExpressions.Regex.Matches(nonHumanToxResposne, pattern,
                System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);

            foreach (System.Text.RegularExpressions.Match match in matchColl)
            {
                string ld50STR = match.Groups[2].Value;
                ld50STR = ld50STR.Replace("&gt;", "");
                ld50STR = ld50STR.Replace("&lt;", "");
                double conc = Convert.ToDouble(ld50STR);
                string unit = match.Groups[3].Value;
                if (string.Compare(unit, "mg/l", true) == 0)
                {
                    if (conc < ld50Dermal)
                    {
                        ld50Dermal = conc;
                        ld50DermalMatch = match;
                    }
                }
                else if (string.Compare(unit, "ug/l", true) == 0)
                {
                    {
                        if (conc / 1000 < ld50Dermal) ld50Dermal = conc / 1000;
                        ld50DermalMatch = match;

                    }
                }
            }

            if (ld50DermalMatch != null)
            {
                m_ld50DermalSpecies = ld50DermalMatch.Groups[1].Value;
                m_ld50DermalValue = ld50Dermal.ToString();
                m_ld50DermalReference = ld50DermalMatch.Groups[4].Value;
            }
        }
    }
}

