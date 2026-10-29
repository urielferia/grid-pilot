using Microsoft.CognitiveServices.Speech;
using Microsoft.CognitiveServices.Speech.Audio;
namespace GridPilot.Providers;

public sealed class SpeechInput : IDisposable
{
    readonly SpeechConfig config;readonly AudioConfig audio;readonly SpeechRecognizer recognizer;readonly List<string> pieces=[];
    CancellationTokenSource? limit;string? failure;
    public SpeechInput(Settings settings,string key)
    {
        if(!settings.CloudEnabled)throw new InvalidOperationException("Active el uso de nube en Ajustes.");
        config=SpeechConfig.FromSubscription(key,settings.SpeechRegion);config.SpeechRecognitionLanguage="es-MX";
        audio=AudioConfig.FromDefaultMicrophoneInput();recognizer=new(config,audio);
        recognizer.Recognized+=(_,e)=>{if(e.Result.Reason==ResultReason.RecognizedSpeech)lock(pieces)pieces.Add(e.Result.Text);};
        recognizer.Canceled+=(_,e)=>failure="Azure Speech canceló el dictado: "+e.Reason;
    }
    public async Task Start(int seconds){pieces.Clear();failure=null;await recognizer.StartContinuousRecognitionAsync();limit=new();_ = StopAfter(Math.Clamp(seconds,1,60),limit.Token);}
    async Task StopAfter(int seconds,CancellationToken token){try{await Task.Delay(TimeSpan.FromSeconds(seconds),token);await recognizer.StopContinuousRecognitionAsync();}catch(OperationCanceledException){} }
    public async Task<string> Stop(){limit?.Cancel();await recognizer.StopContinuousRecognitionAsync();if(failure is not null)throw new InvalidOperationException(failure);lock(pieces)return string.Join(" ",pieces);}
    public void Dispose(){limit?.Cancel();limit?.Dispose();recognizer.Dispose();audio.Dispose();}
}
