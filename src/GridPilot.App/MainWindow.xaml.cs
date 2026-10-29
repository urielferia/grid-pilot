using System.ComponentModel;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Input;
using GridPilot.Core;
using GridPilot.Providers;
using Microsoft.Win32;

namespace GridPilot.App;
public partial class MainWindow : Window
{
    RpcClient? rpc;WorkbookSnapshot? snapshot;WorkbookProfile? profile;bool busy,recording,startingVoice,releaseVoice;SpeechInput? speech;
    readonly string root=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"GridPilot");
    Settings settings=new();readonly List<string> history=[];
    CredentialStore Credentials=>new(Path.Combine(root,"credentials"));
    public MainWindow(){InitializeComponent();Directory.CreateDirectory(root);var file=Path.Combine(root,"settings.json");if(File.Exists(file))settings=Json.Read<Settings>(File.ReadAllText(file));}
    async Task Run(Func<Task> action){if(busy||recording)return;busy=true;Controls.IsEnabled=false;try{await action();}catch(Exception e){Status.Text=e.Message;}finally{busy=false;Controls.IsEnabled=true;}}
    async Task<RpcClient> Client(){if(rpc is null){rpc=new();await rpc.Call("initialize",new{protocolVersion="2025-06-18",capabilities=new{},clientInfo=new{name="GridPilot.UI",version="0.1.0"}});}return rpc;}
    async Task OpenPath(string path)
    {
        Status.Text="Abriendo copia en Excel…";var client=await Client();profile=Json.Read<WorkbookProfile>(File.ReadAllText(Path.Combine(AppContext.BaseDirectory,"assets","profile.json")));
        var result=await client.Call("gridpilot/open",new{path,root=Path.Combine(root,"sessions"),profile});history.Clear();await LoadSnapshot();
        Status.Text=result.GetProperty("Editable").GetBoolean()?"Borrador listo. El original permanece fuera de las escrituras.":"Libro disponible solo para inspección: perfil incompatible.";
    }
    async void Open(object sender,RoutedEventArgs e)=>await Run(async()=>{var dialog=new OpenFileDialog{Filter="Libro Excel|*.xlsx"};if(dialog.ShowDialog()==true)await OpenPath(dialog.FileName);});
    async void Demo(object sender,RoutedEventArgs e)=>await Run(()=>OpenPath(Path.Combine(AppContext.BaseDirectory,"assets","GridPilot-Demo.xlsx")));
    async void Recover(object sender,RoutedEventArgs e)=>await Run(async()=>{var dialog=new OpenFileDialog{Title="Selecciona session.db de la sesión a recuperar",InitialDirectory=Path.Combine(root,"sessions"),Filter="Sesión GridPilot|session.db"};if(dialog.ShowDialog()==true){await(await Client()).Call("gridpilot/recover",new{directory=Path.GetDirectoryName(dialog.FileName)});history.Clear();await LoadSnapshot();Status.Text="Sesión recuperada desde el último estado válido. No se repitieron escrituras.";}});
    async Task LoadSnapshot()
    {
        var r=await(await Client()).Tool("inspect_workbook",new{});snapshot=Json.Read<WorkbookSnapshot>(r.GetProperty("snapshot").GetRawText());profile=Json.Read<WorkbookProfile>(r.GetProperty("profile").GetRawText());
        Structure.Text=string.Join("\n\n",snapshot.Tables.Select(t=>$"{t.Name} · hoja {t.Sheet} · {t.Rows.Count} registros\n{string.Join(", ",t.Headers)}"))+"\n\nRevisión: "+snapshot.Revision;
        var summary=snapshot.Cells.Where(c=>c.Sheet==profile.SummarySheet&&new[]{"B9","B10","B11"}.Contains(c.Address));
        Review.Text=string.Join("\n\n",history)+"\n\nResumen actual: "+string.Join(" · ",summary.Select(c=>$"{c.Address}: {c.Value}"))+"\n\nRevisa también las cuatro hojas en la ventana de Excel. La aprobación cubre la revisión completa.";
    }
    void RequireBook(){if(snapshot is null||profile is null)throw new InvalidOperationException("Primero abre un libro.");}
    async Task Execute(List<Operation> operations,string intent,string revision)
    {
        Status.Text="Modificando y comprobando el borrador…";var plan=new ActionPlan(Guid.NewGuid().ToString(),intent,revision,operations);
        var r=await(await Client()).Tool("execute_batch",new{plan});var change=Json.Read<ChangeSet>(r.GetRawText());
        var text=new StringBuilder(intent+"\n");foreach(var c in change.Changes)text.AppendLine($"{c.Sheet}!{c.Address}: {c.Before?.Value??"∅"} → {c.After?.Value??"∅"}"+(c.Before?.Formula!=c.After?.Formula?$"\nFórmula: {c.Before?.Formula} → {c.After?.Formula}":""));foreach(var s in change.Structure)text.AppendLine(s);history.Add(text.ToString());await LoadSnapshot();Status.Text="Listo para revisar.";
    }
    async void Send(object sender,RoutedEventArgs e)=>await Run(async()=>{
        RequireBook();Status.Text="Interpretando…";await LoadSnapshot();var revision=snapshot!.Revision;var planner=new AnthropicPlanner(settings,Credentials.Read("anthropic"));
        Interpretation result;try{result=await planner.Interpret(Instruction.Text,profile!,async(table,query)=>(object)await(await Client()).Tool("read_records",new{table,query}),CancellationToken.None);}finally{if(planner.LastUsage is not null)await File.AppendAllTextAsync(Path.Combine(root,"usage.jsonl"),Json.Write(planner.LastUsage)+Environment.NewLine);}
        if(result.Clarification!=""){Status.Text="Necesito aclaración: "+result.Clarification;return;}if(result.Operations.Count==0){Status.Text="No se propusieron cambios.";return;}await Execute(result.Operations,Instruction.Text,revision);
    });
    async void ExecuteLocal(object sender,RoutedEventArgs e)=>await Run(async()=>{RequireBook();await Execute(Json.Read<List<Operation>>(Batch.Text),"Lote local",snapshot!.Revision);});
    async void Refresh(object sender,RoutedEventArgs e)=>await Run(async()=>{RequireBook();var before=snapshot!;await LoadSnapshot();if(before.Revision!=snapshot!.Revision){var delta=Diff.Compare(new("","Edición manual",before.Revision,[]),before,snapshot,"");history.Add("Edición manual detectada\n"+string.Join("\n",delta.Changes.Select(c=>$"{c.Sheet}!{c.Address}: {c.Before?.Value} → {c.After?.Value}")));await LoadSnapshot();}Status.Text="Revisión actualizada.";});
    async void Undo(object sender,RoutedEventArgs e)=>await Run(async()=>{RequireBook();await(await Client()).Tool("restore_checkpoint",new{revision=snapshot!.Revision});history.Add("Se deshizo el último lote.");await LoadSnapshot();Status.Text="Borrador restaurado.";});
    async void ApproveExport(object sender,RoutedEventArgs e)=>await Run(async()=>{
        RequireBook();var destination=new SaveFileDialog{Filter="Libro Excel|*.xlsx",FileName="GridPilot-aprobado.xlsx",OverwritePrompt=true};if(destination.ShowDialog()!=true)return;
        if(MessageBox.Show("¿Apruebas la revisión completa mostrada y su exportación a un archivo nuevo?","Aprobar revisión",MessageBoxButton.YesNo,MessageBoxImage.Question)!=MessageBoxResult.Yes)return;
        await(await Client()).Call("gridpilot/approve",new{revision=snapshot!.Revision});var r=await(await Client()).Call("gridpilot/export",new{path=destination.FileName});Status.Text="Archivo aprobado exportado. SHA-256: "+r.GetProperty("Hash").GetString();
    });
    async void Discard(object sender,RoutedEventArgs e)=>await Run(async()=>{await(await Client()).Call("gridpilot/close");snapshot=null;profile=null;history.Clear();Review.Clear();Structure.Clear();Status.Text="Sesión cerrada sin exportar. Los puntos de recuperación locales se conservan.";});
    void Configure(object sender,RoutedEventArgs e)
    {
        var dialog=new SettingsWindow(settings){Owner=this};if(dialog.ShowDialog()!=true)return;settings=dialog.Result;File.WriteAllText(Path.Combine(root,"settings.json"),Json.Write(settings));if(dialog.AnthropicKey!="")Credentials.Save("anthropic",dialog.AnthropicKey);if(dialog.SpeechKey!="")Credentials.Save("speech",dialog.SpeechKey);Status.Text="Ajustes guardados; credenciales protegidas para este usuario de Windows.";
    }
    async void VoiceDown(object sender,MouseButtonEventArgs e)
    {
        if(busy||recording)return;e.Handled=true;recording=true;startingVoice=true;releaseVoice=false;Voice.CaptureMouse();Status.Text="Escuchando…";
        try{speech=new(settings,Credentials.Read("speech"));await speech.Start(settings.MaxVoiceSeconds);}catch(Exception ex){Status.Text=ex.Message;recording=false;speech?.Dispose();speech=null;}finally{startingVoice=false;if(releaseVoice)await FinishVoice();}
    }
    async void VoiceUp(object sender,MouseButtonEventArgs e){e.Handled=true;Voice.ReleaseMouseCapture();await FinishVoice();}
    async void VoiceLost(object sender,MouseEventArgs e)=>await FinishVoice();
    async Task FinishVoice(){if(startingVoice){releaseVoice=true;return;}if(!recording)return;recording=false;busy=true;try{if(speech is not null)Instruction.Text=await speech.Stop();Status.Text="Transcripción lista. Revisa el texto y pulsa Enviar.";}catch(Exception e){Status.Text=e.Message;}finally{speech?.Dispose();speech=null;busy=false;}}
    void OnClosing(object? sender,CancelEventArgs e){if(busy||recording){e.Cancel=true;Status.Text="Espera a que termine la operación o suelta el botón de voz.";return;}speech?.Dispose();rpc?.Dispose();}
}
