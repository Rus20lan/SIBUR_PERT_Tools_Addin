using System;
using Microsoft.Office.Core;
using MSProject = Microsoft.Office.Interop.MSProject;

namespace SIBUR_PERT_Tools_Addin
{
    public partial class ThisAddIn
    {
        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new PertRibbon();
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                dynamic app = this.Application;
                app.WindowSelectionChange += new Action<MSProject.Window>(OnWindowSelectionChange);
                app.ProjectBeforeTaskDelete += new TaskDeleteDelegate(OnProjectBeforeTaskDelete);
            }
            catch (Exception ex)
            {

                System.Diagnostics.Debug.WriteLine($"Ошибка при подписке на события: {ex.Message}");
            }
        }

        private void OnWindowSelectionChange(MSProject.Window window)
        {
            PertManager.Instance.InvalidateRibbon();
        }

        private delegate void TaskDeleteDelegate(MSProject.Task tsk, ref bool cancel);

        private void OnProjectBeforeTaskDelete(MSProject.Task tsk, ref bool cancel)
        {
            PertManager.Instance.InvalidateRibbon();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Код, автоматически созданный VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
