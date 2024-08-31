using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

[
    Guid("339E340A-F510-485B-BBB0-8CD5174E3D8C"),
    ProgId(ServerProgId),
    ComVisible(true)
]
public class SlowAddRtdServer : IRtdServer
{
    class Result
    {
        public DateTime AvailableAfter;
        public double Value;
        public Timer Timer;
    }


    public const string ServerProgId = "ExcelDna.RtdServer.SlowAdd";
    public readonly TimeSpan ResultDelay = TimeSpan.FromSeconds(5);

    private IRTDUpdateEvent m_callback;
    private Dictionary<int, Result> m_topics;
    private bool m_notified;

    public int ServerStart(IRTDUpdateEvent callback)
    {
        Debug.Print($"[{DateTime.Now:HH:mm:ss.fff}] ServerStart");

        m_callback = callback;

        m_topics = new Dictionary<int, Result>();
        m_notified = false;

        var heartbeatInterval = callback.HeartbeatInterval;
        Debug.Print($"[{DateTime.Now:HH:mm:ss.fff}] HeartbeatInterval: {heartbeatInterval}");
        return 1;
    }

    public void ServerTerminate()
    {
        Debug.Print($"[{DateTime.Now:HH:mm:ss.fff}] ServerTerminate");
    }

    public object ConnectData(int topicId,
                              ref Array strings,
                              ref bool newValues)
    {
        var now = DateTime.Now;
        Debug.Print($"[{now:HH:mm:ss.fff}] ConnectData");

        if (2 != strings.Length)
        {
            return "!!! Exactly two parameter are required.";
        }

        double v0 = double.Parse((string)strings.GetValue(0), NumberStyles.Any, CultureInfo.InvariantCulture);
        double v1 = double.Parse((string)strings.GetValue(1), NumberStyles.Any, CultureInfo.InvariantCulture);
        var timer = new Timer
        {
            Interval = (int)ResultDelay.TotalMilliseconds,
        };
        timer.Tick += new EventHandler(TimerEventHandler);

        var result = new Result
        {
            AvailableAfter = now + ResultDelay,
            Value = v0 + v1,
            Timer = timer
        };

        m_topics[topicId] = result;
        
        timer.Start(); // Ensure we'll call UpdateNotify some time later
        return "### Delayed ...";
    }

    public void DisconnectData(int topicId)
    {
        Debug.Print($"[{DateTime.Now:HH:mm:ss.fff}] DisconnectData");
        var result = m_topics[topicId];
        var timer = result.Timer;
        timer.Dispose();
        m_topics.Remove(topicId);
    }

    public Array RefreshData(ref int topicCount)
    {
        var now = DateTime.Now;
        Debug.Print($"[{now:HH:mm:ss.fff}] RefreshData");
        var topics = m_topics.Where(kvp => kvp.Value.AvailableAfter <= now)
                             .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

        object[,] data = new object[2, topics.Count];

        int index = 0;

        foreach (int topicId in topics.Keys)
        {
            var topic = topics[topicId];
            data[0, index] = topicId;
            data[1, index] = topic.Value;

            // No need to keep the timer running
            topic.Timer.Stop();

            ++index;
        }

        topicCount = topics.Count;

        Debug.Print($"[{now:HH:mm:ss.fff}] RefreshData - Returning {topicCount} values: " +
            string.Join(", ", topics.Select(kvp => kvp.Value.Value)));


        Debug.Print($"[{now:HH:mm:ss.fff}] RefreshData - Resetting notified flag");
        m_notified = false;
        return data;
    }

    public int Heartbeat()
    {
        Debug.Print($"[{DateTime.Now:HH:mm:ss.fff}] Heartbeat");
        return 1;
    }

    private void TimerEventHandler(object sender,
                                   EventArgs args)
    {
        var timer = (Timer)sender;
        timer.Stop();
        if (m_notified)
        {
            Debug.Print($"[{DateTime.Now:HH:mm:ss.fff}] Timer Tick - Already Notified");
        }
        else
        {
            Debug.Print($"[{DateTime.Now:HH:mm:ss.fff}] Timer Tick - Calling UpdateNotify and setting notified flag");

            m_callback.UpdateNotify();
            m_notified = true;
        }
    }
}