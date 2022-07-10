using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMWorks
{
    public class TaskChangedBase {
        public delegate void TaskChangedEventHandler(object sender, EventArgs e);
        public event TaskChangedEventHandler TaskChanged;

        public void OnTaskChanged(String taskDesc, DateTime taskStartTime){
            if (TaskChanged != null){
                TaskChangedEventArgs args = new TaskChangedEventArgs();
                args.TaskDesc = taskDesc;
                args.EventTime = taskStartTime;
                TaskChanged(null, args);
            }
        }
    }

    public class TaskChangedEventArgs : EventArgs{
        public String TaskDesc { get; set; }
        public DateTime EventTime { get; set; }
    }
}
