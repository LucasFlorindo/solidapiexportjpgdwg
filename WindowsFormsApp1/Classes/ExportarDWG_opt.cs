using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1.Classes
{
    /// <summary>
    /// Lida com as opções para a geração de um DWG ou DXF.
    /// </summary>
    public enum ExportarDWG_opt
    {
        Exp_FormingTools = 64,
        Exp_LibFiles = 32,
        MergeCoplanarFaces = 16,
        Include_Sketches = 8,
        Exp_BendLines = 4,
        Include_HiddenEdges = 2,
        Exp_Geometry = 1,
    }
}
