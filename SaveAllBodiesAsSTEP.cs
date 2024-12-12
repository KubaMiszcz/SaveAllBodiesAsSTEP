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
        private const string fileExtension = ".STEP";

        public void Main()
        {
            swModel = (ModelDoc2)swApp.ActiveDoc;
            var fullFilePath = swModel.GetPathName();
            //var fullDirectoryPath = Path.GetDirectoryName(fullFilePath);
            var activeConfigurationName = ((IConfiguration)swModel.GetActiveConfiguration()).Name;
            var configurationsCount = swModel.GetConfigurationCount();

            if (IsAssembly_SLDASM(fullFilePath)) // ASSEMBLY
            {
                var componentNames = handleWithAssembly(fullFilePath, activeConfigurationName, configurationsCount);
                ShowSummary(componentNames, "components");
            }

            if (IsPart_SLDPRT(fullFilePath)) // PART
            {
                var bodiesNames = handleWithPart(fullFilePath, activeConfigurationName, configurationsCount);
                ShowSummary(bodiesNames, "bodies");

                //handleWithPartDEPR(fullFilePath, activeConfigurationName, fullDirectoryPath);
            }

            return;
        }

        private List<string> handleWithAssembly(string fullFilePath, string activeConfigurationName, int configurationsCount)
        {
            AssemblyDoc swAssembly = (AssemblyDoc)swModel;

            //no use GetVisibleComponentsInView becuase it gets all in tree
            //and havent switch toplevelonly
            //but i need only toplevel
            var allComponentsInAssembly = ((Object[])swAssembly.GetComponents(ToplevelOnly: true))
                .Select(c => (Component2)c)
                .Where(c => !c.IsSuppressed())
                .ToList();

            allComponentsInAssembly = GetNotForReferenceComponentsOnly(allComponentsInAssembly);
            var processedComponents = new List<Component2>();
            var processedComponentsNames = new List<string>();

            SelectionMgr swSelMgr = (SelectionMgr)swModel.SelectionManager;
            int selectedCount = swSelMgr.GetSelectedObjectCount2(-1);

            if (selectedCount == 0)
            {
                //if nothin slected get all visible components
                processedComponents = allComponentsInAssembly
                    .Where(c => c.Visible == (int)swComponentVisibilityState_e.swComponentVisible)
                    .ToList();
            }
            else
            {
                // if anything selected, save only selected no matter visible or not
                for (int i = 1; i <= selectedCount; i++)
                {
                    // Sprawdzenie, czy zaznaczony obiekt jest typu Component
                    if (swSelMgr.GetSelectedObjectType3(i, -1) == (int)swSelectType_e.swSelCOMPONENTS)
                    {
                        var swComponent = (Component2)swSelMgr.GetSelectedObject6(i, -1);

                        if (swComponent != null)
                        {
                            processedComponents.Add(swComponent);
                        }
                    }
                }
            }


            if (processedComponents.Count > 0)
            {
                try
                {
                    foreach (var component in processedComponents)
                    {
                        Debug.Print("Name of component: " + component.Name2);
                        var filename = Regex.Replace(component.Name2, @"-\d+$", "");

                        if (configurationsCount > 1)
                        {
                            filename += '-' + activeConfigurationName;
                        }

                        component.Select4(false, null, false);

                        SaveToFile(filename, Path.GetDirectoryName(fullFilePath));
                        processedComponentsNames.Add(filename);
                    }
                }
                catch (Exception e)
                {
                    swApp.SendMsgToUser2(e.Message, (int)swMessageBoxIcon_e.swMbStop, (int)swMessageBoxBtn_e.swMbOk);
                    throw;
                }
            }

            return processedComponentsNames;
        }


        private List<string> handleWithPart(string fullFilePath, string activeConfigurationName, int configurationsCount)
        {
            var processedBodiesNames = new List<string>();
            var processedBodies = new List<Body2>();
            PartDoc swPart = (PartDoc)swModel;
            var allBodiesInPart = ((Object[])swPart.GetBodies2((int)swBodyType_e.swSolidBody, BVisibleOnly: true))
                 .Select(b => (Body2)b)
                 .ToList();

            // Pobranie menedżera zaznaczeń
            SelectionMgr swSelMgr = (SelectionMgr)swModel.SelectionManager;

            // Sprawdzenie, czy są zaznaczone obiekty
            int selectedCount = swSelMgr.GetSelectedObjectCount2(-1);

            if (selectedCount == 0)
            {
                //if nothin slected get all visible bodies
                processedBodies = allBodiesInPart.Where(b => b.Visible).ToList();
            }
            else
            {
                // if anything selected, save only selected no matter visible or not
                for (int i = 1; i <= selectedCount; i++)
                {
                    // Sprawdzenie, czy zaznaczony obiekt jest typu Body
                    if (swSelMgr.GetSelectedObjectType3(i, -1) == (int)swSelectType_e.swSelSOLIDBODIES)
                    {
                        var swBody = (Body2)swSelMgr.GetSelectedObject6(i, -1);

                        if (swBody != null)
                        {
                            processedBodies.Add(swBody);
                        }
                    }
                }
            }

            if (processedBodies.Count > 0)
            {
                try
                {
                    var partFileName = Path.GetFileNameWithoutExtension(fullFilePath);

                    foreach (var body in processedBodies)
                    {
                        Debug.Print("Name of body: " + body.Name);

                        var fileName = partFileName;

                        if (configurationsCount > 1)
                        {
                            fileName += '-' + activeConfigurationName;
                        }

                        if (allBodiesInPart.Count > 1)
                        {
                            fileName += ("-" + body.Name);
                            body.Select2(false, null);
                        }

                        SaveToFile(fileName, Path.GetDirectoryName(fullFilePath));
                        processedBodiesNames.Add(fileName);
                    }
                }
                catch (Exception e)
                {
                    swApp.SendMsgToUser2(e.Message, (int)swMessageBoxIcon_e.swMbStop, (int)swMessageBoxBtn_e.swMbOk);
                    throw;
                }
            }

            return processedBodiesNames;
        }

        /// <summary>
        /// ///////////////////////////////////
        /// Privates
        /// ///////////////////////////////////
        /// 

        private bool SaveToFile(string fileName, string fullDirectoryPath)
        {
            var path = Path.Combine(fullDirectoryPath, fileName + fileExtension);
            //result = swModel.SaveAs2(path, 0, true, false);
            //swModel.SaveAs3(path, 0, 2);
            return swModel.SaveAs(path);  //(path, 0, 2);
        }

        private List<Component2> GetNotForReferenceComponentsOnly(List<Component2> components)
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

