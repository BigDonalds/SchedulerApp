using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;


namespace SchedulerApp.Services
{
    public class StateSave
    {
        private const bool ENABLE_SAVING = true;
        private readonly string appDataFolder;
        private readonly string batchesFolder;
        private readonly string schedulesFolder;
        private readonly string employeesFolder;

        public StateSave()
        {
            string appFolder = AppDomain.CurrentDomain.BaseDirectory;
            appDataFolder = Path.Combine(appFolder, "SchedulerData");
            batchesFolder = Path.Combine(appDataFolder, "Batches");
            schedulesFolder = Path.Combine(appDataFolder, "Schedules");
            employeesFolder = Path.Combine(appDataFolder, "Employees");

            if (ENABLE_SAVING)
            {
                Directory.CreateDirectory(batchesFolder);
                Directory.CreateDirectory(schedulesFolder);
                Directory.CreateDirectory(employeesFolder);
            }
        }

        public void SaveEmployee(MainWindow.AvailabilityEntry employee)
        {
            if (!ENABLE_SAVING || employee == null) return;

            try
            {
                string filePath = Path.Combine(employeesFolder, $"{employee.Id}.json");
                string json = JsonConvert.SerializeObject(employee, Newtonsoft.Json.Formatting.Indented);
                File.WriteAllText(filePath, json);
            }
            catch { }
        }

        public void SaveAllEmployees(List<MainWindow.AvailabilityEntry> employees)
        {
            if (!ENABLE_SAVING) return;

            try
            {
                if (Directory.Exists(employeesFolder))
                {
                    var files = Directory.GetFiles(employeesFolder, "*.json");
                    foreach (var file in files) File.Delete(file);
                }

                foreach (var employee in employees) SaveEmployee(employee);
            }
            catch { }
        }

        public void SaveBatch(MainWindow.Batch batch)
        {
            if (!ENABLE_SAVING || batch == null) return;

            try
            {
                string filePath = Path.Combine(batchesFolder, $"{batch.Id}.json");
                string json = JsonConvert.SerializeObject(batch, Newtonsoft.Json.Formatting.Indented);
                File.WriteAllText(filePath, json);
            }
            catch { }
        }

        public void SaveAllBatches(List<MainWindow.Batch> batches)
        {
            if (!ENABLE_SAVING) return;

            try
            {
                if (Directory.Exists(batchesFolder))
                {
                    var files = Directory.GetFiles(batchesFolder, "*.json");
                    foreach (var file in files) File.Delete(file);
                }

                foreach (var batch in batches) SaveBatch(batch);
            }
            catch { }
        }

        public void SaveSchedule(MainWindow.Schedule schedule)
        {
            if (!ENABLE_SAVING || schedule == null) return;

            try
            {
                string filePath = Path.Combine(schedulesFolder, $"{schedule.Id}.json");
                string json = JsonConvert.SerializeObject(schedule, Newtonsoft.Json.Formatting.Indented);
                File.WriteAllText(filePath, json);
            }
            catch { }
        }

        public void SaveAllSchedules(List<MainWindow.Schedule> schedules)
        {
            if (!ENABLE_SAVING) return;

            try
            {
                if (Directory.Exists(schedulesFolder))
                {
                    var files = Directory.GetFiles(schedulesFolder, "*.json");
                    foreach (var file in files) File.Delete(file);
                }

                foreach (var schedule in schedules) SaveSchedule(schedule);
            }
            catch { }
        }

        public List<MainWindow.AvailabilityEntry> LoadEmployees()
        {
            var employees = new List<MainWindow.AvailabilityEntry>();
            if (!ENABLE_SAVING || !Directory.Exists(employeesFolder)) return employees;

            try
            {
                var files = Directory.GetFiles(employeesFolder, "*.json");
                foreach (var file in files)
                {
                    try
                    {
                        string json = File.ReadAllText(file);
                        var employee = JsonConvert.DeserializeObject<MainWindow.AvailabilityEntry>(json);
                        if (employee != null)
                        {
                            if (employee.ScheduleMatrix == null) employee.ScheduleMatrix = new bool[0, 0];
                            employees.Add(employee);
                        }
                    }
                    catch { }
                }
            }
            catch { }
            return employees;
        }

        public List<MainWindow.Batch> LoadBatches()
        {
            var batches = new List<MainWindow.Batch>();
            if (!ENABLE_SAVING || !Directory.Exists(batchesFolder)) return batches;

            try
            {
                var files = Directory.GetFiles(batchesFolder, "*.json");
                foreach (var file in files)
                {
                    try
                    {
                        string json = File.ReadAllText(file);
                        var batch = JsonConvert.DeserializeObject<MainWindow.Batch>(json);
                        if (batch != null) batches.Add(batch);
                    }
                    catch { }
                }
            }
            catch { }
            return batches;
        }

        public List<MainWindow.Schedule> LoadSchedules()
        {
            var schedules = new List<MainWindow.Schedule>();
            if (!ENABLE_SAVING || !Directory.Exists(schedulesFolder)) return schedules;

            try
            {
                var files = Directory.GetFiles(schedulesFolder, "*.json");
                foreach (var file in files)
                {
                    try
                    {
                        string json = File.ReadAllText(file);
                        var schedule = JsonConvert.DeserializeObject<MainWindow.Schedule>(json);
                        if (schedule != null) schedules.Add(schedule);
                    }
                    catch { }
                }
            }
            catch { }
            return schedules;
        }

        public void DeleteEmployee(string employeeId)
        {
            if (!ENABLE_SAVING) return;

            try
            {
                string filePath = Path.Combine(employeesFolder, $"{employeeId}.json");
                if (File.Exists(filePath)) File.Delete(filePath);
            }
            catch { }
        }

        public void DeleteBatch(string batchId)
        {
            if (!ENABLE_SAVING) return;

            try
            {
                string filePath = Path.Combine(batchesFolder, $"{batchId}.json");
                if (File.Exists(filePath)) File.Delete(filePath);
            }
            catch { }
        }

        public void DeleteSchedule(string scheduleId)
        {
            if (!ENABLE_SAVING) return;

            try
            {
                string filePath = Path.Combine(schedulesFolder, $"{scheduleId}.json");
                if (File.Exists(filePath)) File.Delete(filePath);
            }
            catch { }
        }

        public void CleanupOrphanedEmployees(List<MainWindow.Batch> allBatches)
        {
            if (!ENABLE_SAVING) return;

            try
            {
                var referencedEmployeeIds = new HashSet<string>();
                foreach (var batch in allBatches)
                    foreach (var employeeId in batch.EmployeeIds)
                        referencedEmployeeIds.Add(employeeId);

                var employeeFiles = Directory.GetFiles(employeesFolder, "*.json");
                foreach (var file in employeeFiles)
                {
                    string fileName = Path.GetFileNameWithoutExtension(file);
                    if (!referencedEmployeeIds.Contains(fileName)) File.Delete(file);
                }
            }
            catch { }
        }

        public string GetDataFolderPath() => appDataFolder;
        public bool IsSavingEnabled() => ENABLE_SAVING;
    }
}