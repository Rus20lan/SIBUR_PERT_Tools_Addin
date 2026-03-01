using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MSProject = Microsoft.Office.Interop.MSProject;

namespace SIBUR_PERT_Tools_Addin
{
    public class PertManager
    {
        private static readonly Lazy<PertManager> _instance = new Lazy<PertManager>(() => new PertManager());
        public static PertManager Instance => _instance.Value;

        private IRibbonUI _ribbon;
        private const string TableName = "Pert";
        private const string PropName = "PERT_HoursPerDay";

        private PertManager() { }

        public void SetRibbon(IRibbonUI ribbon) => _ribbon = ribbon;

        public void InvalidateRibbon() => _ribbon?.Invalidate();

        private MSProject.Application App => Globals.ThisAddIn.Application;
        private MSProject.Project ActiveProject => App.ActiveProject;

        ///<summary>
        /// Проверка существования таблицы "Pert"
        /// </summary>
        public bool CheckPertTableExists()
        {
            try
            {
                foreach (MSProject.Table tbl in ActiveProject.TaskTables)
                {
                    if (tbl.Name == TableName) return true;
                }
            }
            catch { }
            return false;
        }
        ///<summary>
        /// Создание таблицы PERT по заданному порядку
        /// </summary>
        public void CreatePertTable()
        {
            try
            {
                App.ScreenUpdating = false;
                // 1. Переключаемся на Диаграмму Ганта
                App.ViewApplyEx("Gantt Chart");
                // 2. Создаем/редактируем таблицу
                bool exists = CheckPertTableExists();

                // Последовательное добавление колонок
                EditTableColumn(TableName, "ID", "ID", 5, true);
                EditTableColumn(TableName, "Name", "Название", 25);
                EditTableColumn(TableName, "Duration4", "Optimistic", 15);
                EditTableColumn(TableName, "Duration5", "Most Likely", 15);
                EditTableColumn(TableName, "Duration6", "Pessimistic", 15);
                EditTableColumn(TableName, "Number4", "W1", 5);
                EditTableColumn(TableName, "Number5", "W2", 5);
                EditTableColumn(TableName, "Number6", "W3", 5);
                EditTableColumn(TableName, "Text30", "Status", 15);
                EditTableColumn(TableName, "Duration", "PERT Duration", 15);
                EditTableColumn(TableName, "Duration7", "Расчет длительности PERT", 15);
                EditTableColumn(TableName, "Start", "Start", 12);
                EditTableColumn(TableName, "Finish", "Finish", 12);
                EditTableColumn(TableName, "Predecessors", "Pred", 10);

                // 3. Применяем таблицу
                App.TableApply(TableName);

                // 4. Переименовываем поля (CustomFieldRename)
                App.CustomFieldRename(MSProject.PjCustomField.pjCustomTaskDuration4, "Optimistic Duration");
                App.CustomFieldRename(MSProject.PjCustomField.pjCustomTaskDuration5, "Most Likely Duration");
                App.CustomFieldRename(MSProject.PjCustomField.pjCustomTaskDuration6, "Pessimistic Duration");
                App.CustomFieldRename(MSProject.PjCustomField.pjCustomTaskNumber4, "Optimistic Weight");
                App.CustomFieldRename(MSProject.PjCustomField.pjCustomTaskNumber5, "Most Likely Weight");
                App.CustomFieldRename(MSProject.PjCustomField.pjCustomTaskNumber6, "Pessimistic Weight");
                App.CustomFieldRename(MSProject.PjCustomField.pjCustomTaskText30, "PERT State");
                App.CustomFieldRename(MSProject.PjCustomField.pjCustomTaskDuration7, "PERT Calc Duration");

                InvalidateRibbon();

                MessageBox.Show("Таблица PERT успешно создана!", "SIBUR PERT", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {

                MessageBox.Show($"Ошибка при создании таблицы: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                App.ScreenUpdating = true;
            }
        }

        private void EditTableColumn(string tableName, string fieldName, string title, int width, bool create = false)
        {
            App.TableEditEx(Name: tableName, TaskTable: true, Create: create, FieldName: fieldName, Title: title, Width: width, ShowInMenu: true);
        }

        ///<summary>
        /// Расчет PERT по формуле (O * w1 + M * w2 + P * w3) / (w1+w2+w3)
        /// </summary>
        public void CalculatePERT()
        {
            double hoursPerDay = GetHoursPerDay();
            int minutesPerDay = (int)(hoursPerDay *  60);

            App.ScreenUpdating = false;
            App.Calculation = MSProject.PjCalculation.pjManual;
            foreach (MSProject.Task task in ActiveProject.Tasks)
            {
                if (task == null || Convert.ToBoolean(task.Summary)) continue;

                double opt = Convert.ToDouble(task.GetField(MSProject.PjField.pjTaskDuration4));
                double mostLikely = Convert.ToDouble(task.GetField(MSProject.PjField.pjTaskDuration5));
                double pess = Convert.ToDouble(task.GetField(MSProject.PjField.pjTaskDuration6));

                double w1 = Convert.ToDouble(task.GetField(MSProject.PjField.pjTaskNumber4));
                double w2 = Convert.ToDouble(task.GetField(MSProject.PjField.pjTaskNumber5));
                double w3 = Convert.ToDouble(task.GetField(MSProject.PjField.pjTaskNumber6));

                if (w1 == 0) w1 = 1;
                if (w2 == 0) w2 = 4;
                if (w3 == 0) w3 = 1;

                if (opt > 0 || mostLikely > 0 || pess > 0)
                {
                    double pertMinutes = (opt * w1 + mostLikely * w2 + pess * w3) / (w1 + w2 + w3);

                    double days = pertMinutes / minutesPerDay;
                    double roundedMinutes = Math.Ceiling(days) * minutesPerDay;

                    task.Duration7 = roundedMinutes;
                    task.Text30 = $"Рассчитано ({DateTime.Now:HH:mm})";
                }
            }
            App.Calculation = MSProject.PjCalculation.pjAutomatic;
            App.ScreenUpdating = true;
            MessageBox.Show("Расчет завершен.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        ///<summary>
        /// Применение расчитанной длительности к основной колонке Duration
        /// </summary>
        public void ApplyPERTDurations()
        {
            var result = MessageBox.Show("Вы уверены, что хотите перезаписать основные длительности задач?", "Подтверждение", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result != DialogResult.Yes) return;

            App.ScreenUpdating = false;
            App.Calculation = MSProject.PjCalculation.pjManual;

            foreach (MSProject.Task task in ActiveProject.Tasks)
            {
                if (task == null || Convert.ToBoolean(task.Summary)) continue;

                double calcDuration = Convert.ToDouble(task.GetField(MSProject.PjField.pjTaskDuration7));

                if (calcDuration > 0)
                {
                    task.Duration = calcDuration;
                    task.Text30 = "Применено";
                }
            }
            App.Calculation = MSProject.PjCalculation.pjAutomatic;
            App.ScreenUpdating = true;
        }

        public double GetHoursPerDay()
        {
            try
            {
                var props = (DocumentProperties)ActiveProject.CustomDocumentProperties;
                foreach (DocumentProperty prop in props)
                {
                    if (prop.Name == PropName) return Convert.ToDouble(prop.Value);
                }
            }
            catch { }
            return 8.0;
        }

        public void SetHoursPerDay(string value)
        {
            if (double.TryParse(value, out var hours))
            {
                var props = (DocumentProperties)ActiveProject.CustomDocumentProperties;
                try
                {
                    bool found = false;
                    foreach (DocumentProperty prop in props)
                    {
                        if (prop.Name == PropName)
                        {
                            prop.Value = hours;
                            found = true;
                            break;
                        }
                    }
                    if (!found) props.Add(PropName, false, MsoDocProperties.msoPropertyTypeFloat, hours);
                }
                catch { }
            }
        }
    }
}
