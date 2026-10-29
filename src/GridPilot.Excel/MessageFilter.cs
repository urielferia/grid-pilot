using System.Runtime.InteropServices;
namespace GridPilot.Excel;

[ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IOleMessageFilter
{
    [PreserveSig] int HandleInComingCall(int type,IntPtr caller,int tick,IntPtr info);
    [PreserveSig] int RetryRejectedCall(IntPtr callee,int tick,int reject);
    [PreserveSig] int MessagePending(IntPtr callee,int tick,int pending);
}
internal sealed class MessageFilter : IOleMessageFilter,IDisposable
{
    readonly IOleMessageFilter? previous;
    [DllImport("ole32.dll")]static extern int CoRegisterMessageFilter(IOleMessageFilter? filter,out IOleMessageFilter? previous);
    public MessageFilter()=>CoRegisterMessageFilter(this,out previous);
    public int HandleInComingCall(int type,IntPtr caller,int tick,IntPtr info)=>0;
    public int RetryRejectedCall(IntPtr callee,int tick,int reject)=>reject==2&&tick<5000?100:-1;
    public int MessagePending(IntPtr callee,int tick,int pending)=>2;
    public void Dispose()=>CoRegisterMessageFilter(previous,out _);
}
