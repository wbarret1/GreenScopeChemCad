using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace GreenScopeChemCad
{
    public partial class SelectChemicalForm : Form
    {
        int m_Selected = -1;

        public SelectChemicalForm(SpellAid aid, string desiredCompoundName)
        {
            InitializeComponent();
            this.imageList1.Images.Clear();
            this.listView1.Columns.Add("Name", -2, HorizontalAlignment.Left);
            this.listView1.Columns.Add("CAS Number", -2, HorizontalAlignment.Left);
            this.AddChemicalsToList(aid);
            this.Text = desiredCompoundName + " not found. Please Select the Desired Chemical From the List Below:";
        }

        public string SelectedChemicalName
        {
            get
            {
                if (m_Selected < 0) return string.Empty;
                return ((SpellAidChemical)(this.listView1.Items[m_Selected].Tag)).Name;
            }
        }

        public string SelectedChemicalCAS
        {
            get
            {
                if (m_Selected < 0) return string.Empty;
                return ((SpellAidChemical)(this.listView1.Items[m_Selected].Tag)).CAS;
            }
        }

        void AddChemicalsToList(SpellAid chemicals)
        {
            int i = 0;
            foreach (SpellAidChemical chemical in chemicals.Chemical)
            {
                ListViewItem item = new ListViewItem(chemical.Name, i++);
                item.Tag = chemical;
                item.SubItems.Add(chemical.CAS);
                this.listView1.Items.Add(item);
                this.imageList1.Images.Add(this.PUGGetCompoundImage(chemical.Name, chemical.CAS));
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            m_Selected = -1;
            if (this.listView1.SelectedIndices.Count != 0)
                m_Selected = this.listView1.SelectedIndices[0];
            this.Close();
        }

        Image PUGGetCompoundImage(string compoundName, string casNo)
        {
            string imageReference = "http://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/" + compoundName + "/PNG";
            System.Net.HttpWebRequest request = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(imageReference);
            try
            {
                System.Net.WebResponse response = request.GetResponse();
                return Image.FromStream(response.GetResponseStream());
            }
            catch (System.Exception p_Ex)
            {
                return null;//Properties.Resources.Image1;
            }
        }

        private void listView1_DoubleClick(object sender, EventArgs e)
        {
            m_Selected = -1;
            if (this.listView1.SelectedIndices.Count != 0)
                m_Selected = this.listView1.SelectedIndices[0];
            this.Close();
        }
    }
}
