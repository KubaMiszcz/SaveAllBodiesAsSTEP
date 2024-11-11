using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SaveAllBodiesAsSTEP
{
    public partial class SaveAllBodiesAsSTEP
    {
        // The SldWorks swApp variable is pre-assigned for you.
        public SldWorks swApp;
        public ModelDoc2 swModel;
        private const int Visible = (int)swComponentVisibilityState_e.swComponentVisible;
        private const int NonVisible = (int)swComponentVisibilityState_e.swComponentHidden;
        private const int swThisConfiguration = (int)swInConfigurationOpts_e.swThisConfiguration;
        private const int swSolidBody = (int)swBodyType_e.swSolidBody;
        private const string fileExtension = ".STEP";

        public void Main()
        {
            swModel = (ModelDoc2)swApp.ActiveDoc;
            var fullFilePath = swModel.GetPathName();
            var fullDirectoryPath = Path.GetDirectoryName(fullFilePath);

            
            if (IsAssembly_SLDASM(fullFilePath)) // ASSEMBLY
            {
                AssemblyDoc swAssembly = (AssemblyDoc)swModel;
                var visibleComponents = GetAllVisibleComponents(swAssembly);
                var processedComponents = GetWithoutReferencedComponents(visibleComponents);
                var processedComponentsNames = new List<string>();

                try
                {
                    foreach (var component in processedComponents)
                    {
                        Debug.Print("Name of component: " + component.Name2);
                        component.Select4(false, null, false);
                        var filename = Regex.Replace(component.Name2, @"-\d+$", "");
                        SaveToFile(filename, fullDirectoryPath);
                        processedComponentsNames.Add(filename);
                    }

                    ShowSummary(processedComponentsNames, "components");
                    return;
                }
                catch (Exception e)
                {
                    swApp.SendMsgToUser2(e.Message, (int)swMessageBoxIcon_e.swMbStop, (int)swMessageBoxBtn_e.swMbOk);
                    throw;
                }
            }

            if (IsPart_SLDPRT(fullFilePath)) // PART
            {
                var processedBodiesNames = new List<string>();
                var partFileName = Path.GetFileNameWithoutExtension(fullFilePath);
                PartDoc swPart = (PartDoc)swModel;
                var visibleBodies = ((Object[])swPart.GetBodies2(swSolidBody, true))
                    .Select(b => (Body2)b)
                    .ToList();

                try
                {
                    foreach (var body in visibleBodies)
                    {
                        var fileName = partFileName;
                        Debug.Print("Name of body: " + body.Name);
                        if (visibleBodies.Count > 1)
                        {
                            fileName += ("-" + body.Name);
                            body.Select2(false, null);
                        }

                        SaveToFile(fileName, fullDirectoryPath);
                        processedBodiesNames.Add(fileName);
                    }

                    ShowSummary(processedBodiesNames, "bodies");
                    return;
                }
                catch (Exception e)
                {
                    swApp.SendMsgToUser2(e.Message, (int)swMessageBoxIcon_e.swMbStop, (int)swMessageBoxBtn_e.swMbOk);
                    throw;
                }
            }

            return;
        }




        /// <summary>
        /// ///////////////////////////////////
        /// Privates
        /// ///////////////////////////////////

        private bool SaveToFile(string fileName, string fullDirectoryPath)
        {
            var path = Path.Combine(fullDirectoryPath, fileName + fileExtension);
            //result = swModel.SaveAs2(path, 0, true, false);
            //swModel.SaveAs3(path, 0, 2);
            return swModel.SaveAs(path);  //(path, 0, 2);
        }

        private List<Component2> GetAllVisibleComponents(AssemblyDoc swAssembly)
        {
            //no use GetVisibleComponentsInView becuase it gets all in tree
            //and havent switch toplevelonly
            //but i need only toplevel
            var components = ((Object[])swAssembly.GetComponents(ToplevelOnly: true))
                .Select(c => (Component2)c)
                .Where(c => c.Visible == Visible)
                .Where(c => !c.IsSuppressed())
                .ToList();

            return components;
        }

        private List<Component2> GetWithoutReferencedComponents(List<Component2> components)
        {
            var result = components.Where(c =>
                    !(
                            c.Name2.Trim().ToLower().StartsWith("REF ".ToLower())
                         || c.Name2.Trim().ToLower().EndsWith("REF".ToLower())
                     )
                )
                .ToList();

            return result;
        }

        private void ShowSummary(List<string> processedFileNames, string typeName)
        {
            var msg = "All " + typeName + " exported successfully 👌👌👌\n\n";
            processedFileNames.ToList().ForEach(n => msg += "   ✔️ " + n + fileExtension + "\n");
            swApp.SendMsgToUser2(msg, (int)swMessageBoxIcon_e.swMbInformation, (int)swMessageBoxBtn_e.swMbOk);
        }

        private bool IsAssembly_SLDASM(string fullFilePath)
        {
            return fullFilePath.ToUpper().EndsWith("SLDASM".ToUpper());
        }

        private bool IsPart_SLDPRT(string fullFilePath)
        {
            return fullFilePath.ToUpper().EndsWith("SLDPRT".ToUpper());
        }

    }
}

