using System.Collections.Concurrent;
using System.Runtime.InteropServices;
namespace GridPilot.Excel;

// One dedicated STA owns every RCW. Never retry a mutation after partial execution.
public sealed class StaThread : IDisposable
{
    readonly BlockingCollection<Action> queue = new();
    readonly Thread thread;
    public StaThread() { thread = new Thread(() => { using var filter=new MessageFilter(); foreach(var action in queue.GetConsumingEnumerable()) action(); }) { IsBackground=true, Name="GridPilot Excel STA" }; thread.SetApartmentState(ApartmentState.STA); thread.Start(); }
    public T Run<T>(Func<T> f)
    {
        var completion = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        queue.Add(() => { try { completion.SetResult(f()); } catch(Exception e) { completion.SetException(e); } });
        return completion.Task.GetAwaiter().GetResult();
    }
    public void Run(Action f) => Run(() => { f(); return true; });
    public void Dispose() { queue.CompleteAdding(); thread.Join(TimeSpan.FromSeconds(10)); queue.Dispose(); }
}
internal sealed class ComScope : IDisposable
{
    readonly List<object> values = [];
    public dynamic Own(object value) { if(Marshal.IsComObject(value)) values.Add(value); return value; }
    public void Dispose() { for(int i=values.Count-1;i>=0;i--) if(Marshal.IsComObject(values[i])) Marshal.ReleaseComObject(values[i]); }
}
