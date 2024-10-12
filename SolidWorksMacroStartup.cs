using System;
using System.Text;
using System.Runtime.InteropServices;


namespace Macro2
{
[Guid("fcaaccd5-4cd6-4682-a2d6-5f5c3d34c3d8")]
    public partial class SolidWorksMacro
    {
        // [TODO] Non-user code interferes with exception handling.
        //[System.Diagnostics.DebuggerNonUserCodeAttribute()]

        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)]
        public String Execute (String strDebug)
        {
            String result = "";
            
            if (strDebug != "")
            {
                //
                // In debug mode, do not enclose in try/catch.
                //

                if (this.swApp == null)
                {
                    throw new System.NullReferenceException("SolidWorksMacro.swApp == null");
                }

                this.Main();

                this.swApp = null;
            }
            else
            {
                //
                // In non-debug mode, catch any exceptions, and feed them back to the caller.
                //

                try
                {
                    if (this.swApp == null)
                    {
                        throw new System.NullReferenceException("SolidWorksMacro.swApp == null");
                    }

                    this.Main();
                }
                catch (Exception ex)
                {
                    Exception topex = ex;

                    StringBuilder sb = new StringBuilder();

                    while (ex != null)
                    {
                        sb.AppendLine(String.Format("{0}: {1} (0x{2:X8})", ex.GetType().ToString(), ex.Message, ex.HResult));
                        sb.AppendLine(String.Format("Source = {0}", ex.Source ?? "<empty>"));
                        sb.AppendLine(String.Format("Stack trace = {0}", ex.StackTrace ?? "<empty>"));

                        if (ex.InnerException != null)
                        {
                            sb.AppendLine();
                            sb.AppendLine("Inner exception:");
                        }

                        // Walk the exception chain.
                        ex = ex.InnerException;
                    }
              
                    result = sb.ToString();

                    System.Diagnostics.Debug.WriteLine(result);
                }
                finally
                {
                    this.swApp = null;
                }
            }

            return(result);
        }
    }
}

