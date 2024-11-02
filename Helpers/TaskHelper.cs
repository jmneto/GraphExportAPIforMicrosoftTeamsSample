//MIT License
//
//Copyright (c) 2024 Microsoft - Jose Batista-Neto.
//
//Permission is hereby granted, free of charge, to any person obtaining a copy
//of this software and associated documentation files (the "Software"), to deal
//in the Software without restriction, including without limitation the rights
//to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
//copies of the Software, and to permit persons to whom the Software is
//furnished to do so, subject to the following conditions:
//
//The above copyright notice and this permission notice shall be included in all
//copies or substantial portions of the Software.
//
//THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
//IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
//AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
//LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
//OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
//SOFTWARE.

namespace GraphExportAPIforMicrosoftTeamsSample.Helpers;

// This is a helper class for managing tasks
internal class TaskHelper
{
    // Private Members
    private List<Task> tasks = new List<Task>();
    private readonly int limit = 1;
    private string taskKelperId = string.Empty;

    // Constructor
    public TaskHelper(string taskname, int limit)
    {
        if (limit > 1)
            this.limit = limit;

        this.taskKelperId = taskname;

        LoggerHelper.WriteToConsoleAndLog($"TaskHelper: '{taskKelperId}' Created with task limit {this.limit}");
    }

    // Public Methods

    // Wait for completion and clear all tasks
    // If a task is faulted or cancelled, throw an exception
    public void WaitAndClearTasks()
    {
        lock (tasks)
        {
            try
            {
                if (tasks.Count > 0)
                    LoggerHelper.WriteToConsoleAndLog($"TaskHelper: '{taskKelperId}' Waiting for {tasks.Count} task(s) to complete");

                Task.WaitAll(tasks.ToArray());
            }
            catch (Exception ex)
            {
                throw new Exception($"TaskHelper WaitAndClearTasks:Task Exception {ex.Message}", ex);
            }
            tasks.Clear();

            MonitorHelper.AddTaskInfo(taskKelperId, limit, 0);
        }
    }

    // Add a task and manage the task limits
    // If the limit is reached, wait for any task to complete
    // If a task is faulted or cancelled, throw an exception
    internal void AddTaskAndManageLimit(Task task)
    {
        lock (tasks)
        {
            tasks.Add(task);

            int count = tasks.Count;

            if (count >= limit)
                Task.WaitAny(tasks.ToArray());

            List<Task> removelist = new List<Task>();

            foreach (Task t in tasks)
            {
                if (t.IsFaulted)
                    throw new Exception($"TaskHelper AddTaskAndManageLimit:Task Exception: {t.Exception?.Message}", t.Exception);

                if (t.IsCanceled)
                    throw new Exception($"TaskHelper AddTaskAndManageLimit:Task Cancelled: {t.Exception?.Message}", t.Exception);

                if (t.IsCompleted)
                    removelist.Add(t);
            }

            foreach (Task t in removelist)
                tasks.Remove(t);

            MonitorHelper.AddTaskInfo(taskKelperId, limit, count);
        }
    }
}