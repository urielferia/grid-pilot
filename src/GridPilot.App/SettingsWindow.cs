using System.Windows;
using System.Windows.Controls;
using GridPilot.Providers;
namespace GridPilot.App;
public sealed class SettingsWindow : Window
{
    public Settings Result {get;private set;} public string AnthropicKey=>anthropic.Password;public string SpeechKey=>speech.Password;
    readonly PasswordBox anthropic=new(),speech=new();
    public SettingsWindow(Settings current)
    {
        Result=current;Title="Ajustes de GridPilot";Width=480;Height=550;WindowStartupLocation=WindowStartupLocation.CenterOwner;
        var panel=new StackPanel{Margin=new Thickness(24)};Content=panel;
        var cloud=new CheckBox{Content="Autorizar uso de servicios de nube (puede generar costos)",IsChecked=current.CloudEnabled,Margin=new Thickness(0,0,0,20)};panel.Children.Add(cloud);
        TextBox Field(string title,string value){panel.Children.Add(new TextBlock{Text=title,Margin=new Thickness(0,8,0,4)});var box=new TextBox{Text=value};panel.Children.Add(box);return box;}
        var model=Field("Modelo Anthropic (un cambio requiere reevaluar el piloto)",current.Model);
        panel.Children.Add(new TextBlock{Text="Clave Anthropic (vacío conserva la guardada)",Margin=new Thickness(0,12,0,4)});panel.Children.Add(anthropic);
        var region=Field("Región Azure Speech",current.SpeechRegion);panel.Children.Add(new TextBlock{Text="Clave Azure Speech",Margin=new Thickness(0,12,0,4)});panel.Children.Add(speech);
        var calls=Field("Máximo de llamadas por instrucción (1–10)",current.MaxCalls.ToString());
        var tokens=Field("Máximo de tokens por instrucción",current.MaxTotalTokens.ToString());
        var save=new Button{Content="Guardar",Margin=new Thickness(0,20,0,0)};panel.Children.Add(save);
        save.Click+=(_,_)=>{if(!int.TryParse(calls.Text,out var n)||n<1||n>10||!int.TryParse(tokens.Text,out var t)||t<1000||t>50000||string.IsNullOrWhiteSpace(model.Text)){MessageBox.Show("Revisa los límites y el modelo.");return;}Result=current with{CloudEnabled=cloud.IsChecked==true,Model=model.Text.Trim(),SpeechRegion=region.Text.Trim(),MaxCalls=n,MaxTotalTokens=t};DialogResult=true;};
    }
}
