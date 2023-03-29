using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFormsApp1.Classes;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.IO;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        SWAPI objSW = new SWAPI();
        string nomeArqPrt = @"C:\Users\Lucas Rodrigues\SKA AUTOMACAO DE ENGENHARIAS LTDA\Desenvolvimento Dassault - General\Documentos sobre o time\Integração de colaboradores\Treinamentos\5 - Treinamento API\Códigos fonte\SolidAPI\Peças\Garra.SLDPRT";
        string nomeArqDrw = @"C:\Users\Lucas Rodrigues\SKA AUTOMACAO DE ENGENHARIAS LTDA\Desenvolvimento Dassault - General\Documentos sobre o time\Integração de colaboradores\Treinamentos\5 - Treinamento API\Códigos fonte\SolidAPI\Peças\Garra.SLDDRW";

        public Form1()
        {
            InitializeComponent();
        }

        private void btnAbrirSW_Click(object sender, EventArgs e)
        {

            // Substituir (Visivel.Checked, Versão do seu SW);


            try
            {
                labelsw.Text = "Abrindo SolidWorks...";
                objSW.AbrirSolidworks(Visivel.Checked, 31);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ERRO", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                labelsw.Text = "Pronto!";
            }
        }

        private void btnFecharSW_Click(object sender, EventArgs e)
        {
            try
            {
                labelsw.Text = "Fechando SolidWorks...";
                objSW.FecharSolidworks();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "ERRO" , MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                labelsw.Text = "Pronto!";
            }

        }

        private void btnAbrir_Click(object sender, EventArgs e)
        {
            try
            {
                labelsw.Text = "Abrindo arquivo...";
                objSW.AbrirArquivo(nomeArqDrw, Path.GetExtension(nomeArqDrw));
                MessageBox.Show(objSW.ObtemArquivoAtivo());

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "ERRO", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally
            {
                labelsw.Text = "Pronto!";
            }
        }

        private void btnFechar_Click(object sender, EventArgs e)
        {
            try
            {
                labelsw.Text = "Abrindo arquivo...";
                objSW.FecharArquivo(nomeArqPrt);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "ERRO", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                labelsw.Text = "Pronto!";

            }
        }


    }
}
