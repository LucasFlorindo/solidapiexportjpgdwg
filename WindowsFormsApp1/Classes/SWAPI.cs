using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace WindowsFormsApp1.Classes
{
    class SWAPI
    {
        /// <summary>
        /// Objeto principal da API.
        /// </summary>
        SldWorks swApp = default(SldWorks);


        /// <summary>
        /// ID do processo do SolidWorks.
        /// </summary>
        int ProcessID = 0;


        /// <summary>
        /// Objeto principal para lidar com arquivos
        /// </summary>
        ModelDoc2 swModelDoc = null;

        /// <summary>
        /// Objeto para lidar apenas com peças.
        /// </summary>
        PartDoc swPart = null;

        /// <summary>
        /// Objeto para lidar apenas com montagens.
        /// </summary>
        AssemblyDoc swAsmb = null;

        /// <summary>
        /// Objeto para lidar apenas com Desenhos.
        /// </summary>
        DrawingDoc swDrw = null;

        /// <summary>
        /// Abre o SolidWorks em uma versão específica.
        /// </summary>
        /// <param name="Visivel"></param>
        /// <param name="Versao"></param>

        public void AbrirSolidworks(bool Visivel, int Versao)
        {
            try
            {
                swApp = Activator.CreateInstance(Type.GetTypeFromProgID($"SldWorks.Application.{Versao.ToString()}")) as SldWorks;
                swApp.Visible = Visivel;
                ProcessID = swApp.GetProcessID();



            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao abrir o Solidworks: {ex.Message}\n{ex.StackTrace}");
            }

        }

        public void FecharSolidworks()
        {
            try
            {
                Process[] processos = Process.GetProcesses();
                var sldworks = processos.FirstOrDefault(x => x.Id == ProcessID);
                sldworks.Kill();
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao fechar o Solidworks: {ex.Message}\n{ex.StackTrace}");
            }
        }

        public void AbrirArquivo(string CaminhoArquivo, string Extensao)
        {
            int err = 0, wars = 0;

            // swDoocumentTypes_e. o ponto puxa qual tipo de documento irá abrir. Nesse exemplo vou abrir um documento de uma peça.
            // essas invocações 'int' são requisitos da função swApp.Opendoc.
            // a configuração "", representa a execução no vazio, default

            try
            {

                if (Extensao.ToUpper() == ".SLDPRT")
                {
                    swModelDoc = swApp.OpenDoc6(CaminhoArquivo, (int)swDocumentTypes_e.swDocPART, 0, "", err, wars);
                    swPart = (PartDoc)swModelDoc;
                }
                else if (Extensao.ToUpper() == ".SLDASM")
                {
                    swModelDoc = swApp.OpenDoc6(CaminhoArquivo, (int)swDocumentTypes_e.swDocASSEMBLY, 0, "", err, wars);
                    swAsmb = (AssemblyDoc)swModelDoc;
                }
                else if (Extensao.ToUpper() == ".SLDDRW")
                {
                    swModelDoc = swApp.OpenDoc6(CaminhoArquivo, (int)swDocumentTypes_e.swDocDRAWING, 0, "", err, wars);
                    swDrw = (DrawingDoc)swModelDoc;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao abrir um arquivo no Solidworks: {ex.Message}\n{ex.StackTrace}");
            }
        }

        public void FecharArquivo(string CaminhoArquivo)
        {
            try
            {
                swApp.CloseDoc(CaminhoArquivo);
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao fechar um arquivo no Solidworks: {ex.Message}\n{ex.StackTrace}");
            }
        }


        //Busca o arquivo ativo no solidworks. "Mais recente"
        public string ObtemArquivoAtivo()
        {
            string arquivo = "";
            try
            {
                swModelDoc = (ModelDoc2)swApp.ActiveDoc;
                arquivo = swModelDoc.GetTitle();
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao fechar um arquivo no Solidworks: {ex.Message}\n{ex.StackTrace}");
            }

            return arquivo;
        }

        // verify if the archive has sheetmetal resource.
        private bool VerificaSheetMetal()
        {
            // recebe a primeira feature do arquivo
            Feature feat = (Feature)swModelDoc.FirstFeature();

            // inicia um LOOP para percorrer todas as features.
            while (feat != null)
            {
                // se encontrar uma feature  om sheetmetal retorna true.
                if (feat.GetTypeName2().ToUpper() == "SHEETMETAL")
                {
                    return true;
                }
                feat = (Feature)feat.GetNextFeature();
            }
            return false;
        }



        public void ExportarParaJPG(string PathDestino)
        {
            try
            {
                bool resultado = swModelDoc.SaveAs(PathDestino);

                if (!resultado)
                {
                    throw new ArgumentException($"Erro em {nameof(ExportarParaJPG)}: Não foi possível converter o arquivo para JPG.");
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro em {nameof(ExportarParaJPG)}: Não foi possível converter o arquivo para JPG. \n{ex.Message}\n{ex.StackTrace}");
            }
        }

        public void ExportarParaDWG(string CaminhoOrigem, string PathDestino)
        {
            try
            {
                if (VerificaSheetMetal())
                {
                    //Here we define whats gonna be exported into DWG.
                    int Opcoes = (int)ExportarDWG.Exp_Geometry + (int)ExportarDWG.Exp_LibFiles + (int)ExportarDWG.Exp_BendLines;

                    bool resultado = swPart.ExportToDWG2(PathDestino, CaminhoOrigem, (int)swExportToDWG_e.swExportToDWG_ExportSheetMetal, true, null, false, false, Opcoes, null);

                    if (!resultado)
                    {
                        throw new Exception($"Erro em {nameof(ExportarParaDWG)}: Não foi possível converter o arquivo para DWG");
                    }
                    else
                    {
                        throw new Exception($"Erro em {nameof(ExportarParaDWG)}: Este arquivo não é uma chapa metálica.");
                    }

                 

                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro em {nameof(ExportarParaDWG)}: Não foi possível converter o arquivo para DWG");
            }
        }

    }
}
