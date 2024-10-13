using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;

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

        public void Main()
        {
            swModel = (ModelDoc2)swApp.ActiveDoc;
            var fullFilePath = swModel.GetPathName();
            var fullDirectoryPath = Path.GetDirectoryName(fullFilePath);

            if (IsAssembly_SLDASM(fullFilePath)) // ASSEMBLY
            {
                AssemblyDoc swAssembly = (AssemblyDoc)swModel;
                //swModel.ViewDisplayWireframe();

                var visibleComponents = GetAllVisibleComponents(swAssembly);
                var processedComponents = RemoveReferenceOnlyComponents(visibleComponents);
                var visibleComponentsNames = processedComponents.Select(c => c.Name2);

                TemporaryHideAllJustForSaving(processedComponents);
                foreach (var component in processedComponents)
                {
                    Debug.Print("Name of component: " + component.Name2);
                    ShowComponent(component);
                    var filename = Regex.Replace(component.Name2, @"-\d+$", "");
                    SaveFile(filename, fullDirectoryPath);
                    HideComponent(component);
                }

                RestoreVisibility(processedComponents);
                //swModel.ViewDisplayShaded();

                return;
            }

            if (IsPart_SLDPRT(fullFilePath)) // PART
            {
                var fileName = Path.GetFileNameWithoutExtension(fullFilePath);
                PartDoc swPart = (PartDoc)swModel;
                var visibleBodies = ((Object[])swPart.GetBodies2(swSolidBody, true))
                    .Select(b => (Body2)b)
                    .ToList();

                var visibleBodiesNames = visibleBodies.Select(b => b.Name);

                foreach (var body in visibleBodies)
                {
                    Debug.Print("Name of body: " + body.Name);
                    if (visibleBodies.Count > 1)
                    {
                        SaveFile(fileName + "-" + body.Name, fullDirectoryPath);
                    }
                    else
                    {
                        SaveFile(fileName, fullDirectoryPath);
                    }
                }

                return;
            }

            return;
        }

        private List<Component2> RemoveReferenceOnlyComponents(List<Component2> visibleComponents)
        {
            return visibleComponents.Where(c =>
            c.Name2.ToLower().StartsWith("REF ".ToLower())
            || c.Name2.ToLower().EndsWith("REF".ToLower())
            ).ToList();
        }


        /// <summary>
        /// ///////////////////////////////////
        /// Privates
        /// ///////////////////////////////////

        private void ShowComponent(Component2 component)
        {
            //component.SetVisibility(Visible, swThisConfiguration, null);
            component.Visible = Visible;
        }

        private void HideComponent(Component2 component)
        {
            //component.SetVisibility(NonVisible, swThisConfiguration, null);
            component.Visible = NonVisible;
        }

        private void TemporaryHideAllJustForSaving(List<Component2> visibleComponents)
        {
            visibleComponents.ForEach(c => HideComponent(c));
        }

        private void RestoreVisibility(List<Component2> visibleComponents)
        {
            visibleComponents.ForEach(c => ShowComponent(c));
        }

        private bool SaveFile(string fileName, string fullDirectoryPath)
        {
            var path = Path.Combine(fullDirectoryPath, fileName + ".STEP");
            //result = swModel.SaveAs2(path, 0, true, false);
            //swModel.SaveAs3(path, 0, 2);
            return swModel.SaveAs(path);  //(path, 0, 2);
        }

        private List<Component2> GetAllVisibleComponents(AssemblyDoc swAssembly)
        {
            //no use GetVisibleComponentsInView becuase it gets all in tree
            //and havent switch toplevelonly
            //but i need only toplevel
            var allComponents = ((Object[])swAssembly.GetComponents(ToplevelOnly: true))
                .Select(c => (Component2)c)
                .Where(c => c.Visible == Visible)
                .ToList();

            return allComponents;
        }

        private bool IsAssembly_SLDASM(string fullFilePath)
        {
            return fullFilePath.ToUpper().EndsWith("SLDASM".ToUpper());
        }

        private bool IsPart_SLDPRT(string fullFilePath)
        {
            return fullFilePath.ToUpper().EndsWith("SLDPRT".ToUpper());
        }



        //private void Main2()
        //{

        //    ModelDoc2 swDoc = null;
        //    PartDoc swPart = null;
        //    DrawingDoc swDrawing = null;
        //    AssemblyDoc swAssembly = null;
        //    bool boolstatus = false;
        //    int longstatus = 0;
        //    int longwarnings = 0;
        //    swDoc = ((ModelDoc2)(swApp.ActiveDoc));
        //    boolstatus = swDoc.Extension.SelectByRay(-0.024179604907430985D, -0.0051289340948983408D, 0.048000000625847861D, -0.8236958650258458D, -0.189895394735217D, -0.53428911742396534D, 0.00020962654891361061D, 2, false, 0, 0);
        //    // 
        //    // Save As
        //    longstatus = swDoc.SaveAs3("D:\\macrotest-skasuj\\TEST =ASM TTGO T-Beam V1.2 Case TEST  antenna inside.STEP", 0, 2);


        //    return;
        //}

    }
}

