//----------------------------------------------------------------------
// <copyright file="SpeechSynthesizer.cs" company="Alisson Sol">
//   Code provided "as is", with full rights for any use or change.
//   References
//     https://docs.microsoft.com/en-us/azure/cognitive-services/speech-service/quickstarts/setup-platform
//     https://azure.microsoft.com/en-us/services/cognitive-services/text-to-speech/#features
//     https://github.com/Azure-Samples/cognitive-services-speech-sdk/blob/master/quickstart/csharp/dotnetcore/text-to-speech/helloworld/Program.cs
// </copyright>
//----------------------------------------------------------------------

using System;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using Microsoft.CognitiveServices.Speech;

namespace PptPolly
{
    internal class SpeechSynthesizerHelper
    {
        /// <summary>
        /// Private constructor to prevent generation of default constructor.
        /// </summary>
        private SpeechSynthesizerHelper()
        {
        }

        static public string GeneratedFilename;

        static public async Task SynthesisToSpeakerAsync()
        {
            // Creates an instance of a speech config with specified subscription key and service region.
            // Replace with your own subscription key and service region (e.g., "westus").
            var config = SpeechConfig.FromSubscription("YourSubscriptionKey", "westus");

            // Creates a speech synthesizer using the default speaker as audio output.
            using (var synthesizer = new SpeechSynthesizer(config))
            {
                // Receive a text from console input and synthesize it to speaker.
                Console.WriteLine("Type some text that you want to speak...");
                Console.Write("> ");
                string text = Console.ReadLine();

                using (var result = await synthesizer.SpeakTextAsync(text))
                {
                    if (result.Reason == ResultReason.SynthesizingAudioCompleted)
                    {
                        Console.WriteLine($"Speech synthesized to speaker for text [{text}]");

                        GeneratedFilename = Path.Combine(Path.GetTempPath(), DateTime.Now.Ticks.ToString(CultureInfo.InvariantCulture) + "_" + Guid.NewGuid().ToString() + ".mp3");
                        using (var fileStream = File.Create(GeneratedFilename))
                        {
                            fileStream.Write(result.AudioData, 0, result.AudioData.Length);
                            fileStream.Flush();
                        }
                    }
                    else if (result.Reason == ResultReason.Canceled)
                    {
                        var cancellation = SpeechSynthesisCancellationDetails.FromResult(result);
                        Console.WriteLine($"CANCELED: Reason={cancellation.Reason}");

                        if (cancellation.Reason == CancellationReason.Error)
                        {
                            Console.WriteLine($"CANCELED: ErrorCode={cancellation.ErrorCode}");
                            Console.WriteLine($"CANCELED: ErrorDetails=[{cancellation.ErrorDetails}]");
                            Console.WriteLine($"CANCELED: Did you update the subscription info?");
                        }
                    }
                }
            }
        }
    }
}
