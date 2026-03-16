using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using MSProject = Microsoft.Office.Interop.MSProject;
using System.Drawing;
//using Microsoft.Office.Interop.MSProject;
using Exception = System.Exception;

namespace SIBUR_PERT_Tools_Addin
{
    public class PertManager
    {
        private static readonly Lazy<PertManager> _instance = new Lazy<PertManager>(() => new PertManager());
        public static PertManager Instance => _instance.Value;

        private const string ViewName = "SIBUR PERT View";

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
                if (ActiveProject == null) return;
                App.FilterClear();
                string defaultFilter = ActiveProject.CurrentFilter;
                App.GroupClear();
                string defaultGroup = ActiveProject.CurrentGroup;

                // Останавливаем отрисовку, чтобы не было "дерганий"
                App.ScreenUpdating = false;
                // 1. Создаем таблицу
                bool tableExists = CheckPertTableExists();
                // Добавлеем в текущий проект новую таблицу Pert
                ActiveProject.TaskTables.Add(TableName, MSProject.PjField.pjTaskID);
                MSProject.Table newTable = ActiveProject.TaskTables[TableName];
                // Последовательное добавление колонок
                newTable.TableFields.Add(MSProject.PjField.pjTaskName,Title: "Название", Width:25,AlignData:MSProject.PjAlignment.pjLeft);
                newTable.TableFields.Add(MSProject.PjField.pjTaskDuration4, Title: "Optimistic", Width: 15);
                newTable.TableFields.Add(MSProject.PjField.pjTaskDuration5, Title: "Most Likely", Width: 15);
                newTable.TableFields.Add(MSProject.PjField.pjTaskDuration6, Title: "Pessimistic", Width: 15);
                newTable.TableFields.Add(MSProject.PjField.pjTaskNumber4, Title: "W1", Width: 5);
                newTable.TableFields.Add(MSProject.PjField.pjTaskNumber5, Title: "W2", Width: 5);
                newTable.TableFields.Add(MSProject.PjField.pjTaskNumber6, Title: "W3", Width: 5);
                newTable.TableFields.Add(MSProject.PjField.pjTaskText30, Title: "Status", Width: 15);
                newTable.TableFields.Add(MSProject.PjField.pjTaskDuration, Title: "Длительность", Width: 15);
                newTable.TableFields.Add(MSProject.PjField.pjTaskDuration7, Title: "Расчет длительности PERT", Width: 15);
                newTable.TableFields.Add(MSProject.PjField.pjTaskStart, Title: "Начало", Width: 12);
                newTable.TableFields.Add(MSProject.PjField.pjTaskFinish, Title: "Окончание", Width: 12);
                newTable.TableFields.Add(MSProject.PjField.pjTaskPredecessors, Width: 10);
                // Создаем представление
                bool viewExists = false;
                foreach (MSProject.View view in ActiveProject.Views)
                {
                    if (view.Name == ViewName) { viewExists = true; break; }
                }
                App.ViewEditSingle(
                   Name: ViewName,
                   Create: !viewExists,
                   Screen: MSProject.PjViewScreen.pjGantt,
                   Table: TableName,
                   Filter: defaultFilter,
                   Group: defaultGroup
                );
                // Применяем результат
                App.ViewApply(ViewName);
                //App.SelectBeginning();
                InvalidateRibbon();
                MessageBox.Show($"Представление '{ViewName}' создано и выбрано.", "SIBUR PERT", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                
                MessageBox.Show($"Ошибка при настройке интерфейса: {ex.Message}\nСтрока:\n{ex.StackTrace}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                App.ScreenUpdating = true;
            }
        }


        //public void CreatePertTable()
        //{
        //    try
        //    {
        //        if (ActiveProject == null) return;
        //        App.ScreenUpdating = false;
        //        // 1. Переключаемся на Диаграмму Ганта
        //        // Динамический поиск представления Диаграммы Ганта для текущей локализации
        //        string ganttViewName = "";
        //        foreach (MSProject.View view in ActiveProject.Views)
        //        {
        //            try
        //            {
        //                if (view != null && view.Screen == MSProject.PjViewScreen.pjGantt)
        //                {
        //                    ganttViewName = view.Name;
        //                    break;
        //                }
        //            }
        //            catch (Exception)
        //            {
        //                continue;
        //            }
                    
        //        }

        //        if (string.IsNullOrEmpty(ganttViewName))
        //        {
        //            try { App.ViewApplyEx(Name: "Диаграмма Ганта"); }
        //            catch { App.ViewApplyEx(Name: "Gantt Chart"); }
        //        }
        //        else
        //        {
        //            App.ViewApplyEx(Name: ganttViewName);
        //        }
        //        // 2. Создаем/редактируем таблицу
        //        bool exists = CheckPertTableExists();

        //        // Последовательное добавление колонок
        //        EditTableColumn(TableName, "ID", "ID", 5, true);
        //        EditTableColumn(TableName, "Name", "Название", 25);
        //        EditTableColumn(TableName, "Duration4", "Optimistic", 15);
        //        EditTableColumn(TableName, "Duration5", "Most Likely", 15);
        //        EditTableColumn(TableName, "Duration6", "Pessimistic", 15);
        //        EditTableColumn(TableName, "Number4", "W1", 5);
        //        EditTableColumn(TableName, "Number5", "W2", 5);
        //        EditTableColumn(TableName, "Number6", "W3", 5);
        //        EditTableColumn(TableName, "Text30", "Status", 15);
        //        EditTableColumn(TableName, "Duration", "PERT Duration", 15);
        //        EditTableColumn(TableName, "Duration7", "Расчет длительности PERT", 15);
        //        EditTableColumn(TableName, "Start", "Start", 12);
        //        EditTableColumn(TableName, "Finish", "Finish", 12);
        //        //EditTableColumn(TableName, "Predecessors", "Pred", 10);
        //        EditTableColumn(TableName, App.FieldConstantToFieldName(MSProject.PjField.pjTaskPredecessors), "Pred", 10);

        //        // 3. Применяем таблицу
        //        App.TableApply(TableName);

        //        // 4. Переименовываем поля (CustomFieldRename)
        //        App.CustomFieldRename(MSProject.PjCustomField.pjCustomTaskDuration4, "Optimistic Duration");
        //        App.CustomFieldRename(MSProject.PjCustomField.pjCustomTaskDuration5, "Most Likely Duration");
        //        App.CustomFieldRename(MSProject.PjCustomField.pjCustomTaskDuration6, "Pessimistic Duration");
        //        App.CustomFieldRename(MSProject.PjCustomField.pjCustomTaskNumber4, "Optimistic Weight");
        //        App.CustomFieldRename(MSProject.PjCustomField.pjCustomTaskNumber5, "Most Likely Weight");
        //        App.CustomFieldRename(MSProject.PjCustomField.pjCustomTaskNumber6, "Pessimistic Weight");
        //        App.CustomFieldRename(MSProject.PjCustomField.pjCustomTaskText30, "PERT State");
        //        App.CustomFieldRename(MSProject.PjCustomField.pjCustomTaskDuration7, "PERT Calc Duration");

        //        InvalidateRibbon();

        //        MessageBox.Show("Таблица PERT успешно создана!", "SIBUR PERT", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //    }
        //    catch (Exception ex)
        //    {

        //        MessageBox.Show($"Ошибка при создании таблицы: {ex.Message}\nСтрока:\n{ex.StackTrace}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //    finally
        //    {
        //        App.ScreenUpdating = true;
        //    }
        //}

        private void EditTableColumn(string tableName, string fieldName, string title, int width, bool create = false)
        {
            App.TableEditEx(
                Name: tableName, 
                TaskTable: true, 
                Create: true, // Не создаем новую, только редактируем
                OverwriteExisting: false,// Не перезаписываем существующие, добавляем новые
                FieldName: fieldName, 
                Title: title, 
                Width: width,
                Align: 1, // pjLeft - выравниваем слева
                ColumnPosition: -1,// Добавляем в конец таблицы
                ShowInMenu: true
            );
        }

        ///<summary>
        /// Расчет PERT по формуле (O * w1 + M * w2 + P * w3) / (w1+w2+w3)
        /// </summary>
        public void CalculatePERT()
        {
            double hoursPerDay = GetHoursPerDay();
            int minutesPerDay = (int)(hoursPerDay *  60);

            if (!ValidateWeights()) return;

            App.ScreenUpdating = false;
            App.Calculation = MSProject.PjCalculation.pjManual;
            foreach (MSProject.Task task in ActiveProject.Tasks)
            {
                if (task == null || Convert.ToBoolean(task.Summary)) continue;
                double opt = Convert.ToDouble(task.Duration4);
                double mostLikely = Convert.ToDouble(task.Duration5);
                double pess = Convert.ToDouble(task.Duration6);
                
                double w1 = Convert.ToDouble(task.Number4);
                double w2 = Convert.ToDouble(task.Number5);
                double w3 = Convert.ToDouble(task.Number6);
                

                
                if (w1 == 0) { w1 = 1; task.Number4 = 1; } 
                if (w2 == 0) { w2 = 4; task.Number5 = 4; }
                if (w3 == 0) { w3 = 1; task.Number6 = 1; }

                Debug.WriteLine($"w1 = {w1}, w2 = {w2}, w3 = {w3}");

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
        /// Эта функция только проверяет и красит. Возвращает false, если есть ошибки
        /// </summary>
        private bool ValidateWeights()
        {
            try
            {
                if(ActiveProject == null) return false;

                App.ScreenUpdating = false;
                int salmonColor = ColorTranslator.ToOle(ColorTranslator.FromHtml("#fa786e"));
                string w1Col = App.FieldConstantToFieldName(MSProject.PjField.pjTaskNumber4);
                string w2Col = App.FieldConstantToFieldName(MSProject.PjField.pjTaskNumber5);
                string w3Col = App.FieldConstantToFieldName(MSProject.PjField.pjTaskNumber6);

                bool hasErrors = false;

                App.OutlineShowAllTasks();

                foreach (MSProject.Task task in ActiveProject.Tasks) 
                {
                    if (task == null || Convert.ToBoolean(task.Summary)) continue;

                    double w1 = task.Number4;
                    double w2 = task.Number5;
                    double w3 = task.Number6;

                    App.SelectTaskField(Row: task.ID, Column: w1Col, RowRelative: false, Width: 2);
                    App.Font32Ex(CellColor: 0xFFFFFF);

                    if ((w1 + w2 + w3) != 6 || w1< 0 || w2 < 0 || w3 <0) 
                    {

                        hasErrors = true;
                        App.SelectTaskField(Row: task.ID, Column: w1Col, RowRelative: false, Width:2);
                        App.Font32Ex(CellColor: salmonColor);
                    }
                }

                App.ScreenUpdating = true;

                if (hasErrors) 
                {
                    MessageBox.Show(
                            "Обнаружены ошибки в весах задач (W1, W2, W3).\n\n" +
                            "По правилу PERT сумма весов должна быть равна 6 (стандарт: 1, 4, 1).\n"+
                            "Проблемные ячейки подсвечены красным цветом. Пожалуйста, исправьте их и нажмите 'Расчет' снова.",
                            "Ошибка валидации",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning
                        );
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                App.ScreenUpdating = true;
                MessageBox.Show("Ошибка при проверке весов! " + ex.Message);
                return false;

            }
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

                double calcDuration = Convert.ToDouble(task.Duration7);

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
