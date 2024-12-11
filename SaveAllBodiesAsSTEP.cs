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
            var fullDirectoryPath = Path.GetDirectoryName(fullFilePath);
            var activeConfigurationName = ((IConfiguration)swModel.GetActiveConfiguration()).Name;
            var configurationsCount = swModel.GetConfigurationCount();

            if (IsAssembly_SLDASM(fullFilePath)) // ASSEMBLY
            {
                handleWithAssembly(fullDirectoryPath);
            }

            if (IsPart_SLDPRT(fullFilePath)) // PART
            {
                handleWithPart(fullFilePath, activeConfigurationName, fullDirectoryPath, configurationsCount);
                //handleWithPartDEPR(fullFilePath, activeConfigurationName, fullDirectoryPath);
            }

            return;
        }

        private void handleWithAssembly(string fullDirectoryPath)
        {
            AssemblyDoc swAssembly = (AssemblyDoc)swModel;
            var visibleComponents = GetAllVisibleComponentsInAssembly(swAssembly);
            var processedComponents = GetNotForReferenceComponentsOnly(visibleComponents);
            var processedComponentsNames = new List<string>();

            //TODO:
            //check if anything slelected and save only selected no matter isible or not
            //if nothin slected get all visible components and save

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


        private void handleWithPart(string fullFilePath, string activeConfigurationName, string fullDirectoryPath, int configurationsCount)
        {
            var processedBodiesNames = new List<string>();
            var processedBodies = new List<Body2>();
            var partFileName = Path.GetFileNameWithoutExtension(fullFilePath);
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
                //if nothin slected get all visible bodies and save
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
        }

        /// <summary>
        /// ///////////////////////////////////
        /// Privates
        /// ///////////////////////////////////
        /// 

        private void handleWithPartDEPR(string fullFilePath, string activeConfigurationName, string fullDirectoryPath)
        {
            var processedBodiesNames = new List<string>();
            var partFileName = Path.GetFileNameWithoutExtension(fullFilePath);
            PartDoc swPart = (PartDoc)swModel;

            var visibleBodies = ((Object[])swPart.GetBodies2((int)swBodyType_e.swSolidBody, true))
                .Select(b => (Body2)b) //it gets also nonvisible body? why?
                                       //.Where(b=>b.Visible) //
                .ToList();

            try
            {
                foreach (var body in visibleBodies)
                {
                    var fileName = partFileName + '-' + activeConfigurationName;

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



        private bool SaveToFile(string fileName, string fullDirectoryPath)
        {
            var path = Path.Combine(fullDirectoryPath, fileName + fileExtension);
            //result = swModel.SaveAs2(path, 0, true, false);
            //swModel.SaveAs3(path, 0, 2);
            return swModel.SaveAs(path);  //(path, 0, 2);
        }

        private List<Component2> GetAllVisibleComponentsInAssembly(AssemblyDoc swAssembly)
        {
            //no use GetVisibleComponentsInView becuase it gets all in tree
            //and havent switch toplevelonly
            //but i need only toplevel

            const int Visible = (int)swComponentVisibilityState_e.swComponentVisible;
            var components = ((Object[])swAssembly.GetComponents(ToplevelOnly: true))
                .Select(c => (Component2)c)
                .Where(c => c.Visible == Visible)
                .Where(c => !c.IsSuppressed())
                .ToList();

            return components;
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

