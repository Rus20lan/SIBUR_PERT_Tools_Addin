using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;

// TODO:  Выполните эти шаги, чтобы активировать элемент XML ленты:

// 1: Скопируйте следующий блок кода в класс ThisAddin, ThisWorkbook или ThisDocument.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new PertRibbon();
//  }

// 2. Создайте методы обратного вызова в области "Обратные вызовы ленты" этого класса, чтобы обрабатывать действия
//    пользователя, например нажатие кнопки. Примечание: если эта лента экспортирована из конструктора ленты,
//    переместите свой код из обработчиков событий в методы обратного вызова и модифицируйте этот код, чтобы работать с
//    моделью программирования расширения ленты (RibbonX).

// 3. Назначьте атрибуты тегам элементов управления в XML-файле ленты, чтобы идентифицировать соответствующие методы обратного вызова в своем коде.  

// Дополнительные сведения можно найти в XML-документации для ленты в справке набора средств Visual Studio для Office.


namespace SIBUR_PERT_Tools_Addin
{
    [ComVisible(true)]
    public class PertRibbon : Office.IRibbonExtensibility
    {
        private IRibbonUI ribbon;

        #region Элементы IRibbonExtensibility

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("SIBUR_PERT_Tools_Addin.PertRibbon.xml");
        }

        public void OnRibbonLoad(IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            PertManager.Instance.SetRibbon(ribbonUI)
        }

        #endregion

        #region Обратные вызовы ленты
        //Информацию о методах создания обратного вызова см. здесь. Дополнительные сведения о методах добавления обратного вызова см. по ссылке https://go.microsoft.com/fwlink/?LinkID=271226
        
        // Callbacks для кнопок
        public void CreatePertTable(IRibbonControl control) => PertManager.Instance.CreatePertTable();
        public void CalculatePERT(IRibbonControl control) => PertManager.Instance.CalculatePERT();
        public void ApplyPERTDurations(IRibbonControl control) => PertManager.Instance.ApplyPERTDurations();
        #endregion


        // Callbacks для состояния кнопок
        public bool GetCreateTableEnabled(IRibbonControl control) => !PertManager.Instance.CheckPertTableExists();
        public bool GetOtherButtonsEnabled(IRibbonControl control) => PertManager.Instance.CheckPertTableExists();

        // Callbacks для EditBox
        public string GetHoursPerDayText(IRibbonControl control) => PertManager.Instance.GetHoursPerDay().ToString();
        public void HoursPerDay_Changed(IRibbonControl control, string text) => PertManager.Instance.SetHoursPerDay(text);


        #region Вспомогательные методы

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
