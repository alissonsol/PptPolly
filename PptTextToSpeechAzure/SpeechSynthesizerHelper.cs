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
using System.Threading;
using System.Threading.Tasks;
using Microsoft.CognitiveServices.Speech;
using Microsoft.CognitiveServices.Speech.Audio;

namespace PptTextToSpeechAzure
{
    public static class AsyncHelper
    {
        private static readonly TaskFactory _taskFactory = new
            TaskFactory(CancellationToken.None,
                        TaskCreationOptions.None,
                        TaskContinuationOptions.None,
                        TaskScheduler.Default);

        public static TResult RunSync<TResult>(Func<Task<TResult>> func)
            => _taskFactory
                .StartNew(func)
                .Unwrap()
                .GetAwaiter()
                .GetResult();

        public static void RunSync(Func<Task> func)
            => _taskFactory
                .StartNew(func)
                .Unwrap()
                .GetAwaiter()
                .GetResult();
    }

    internal class SpeechSynthesizerHelper
    {
        /// <summary>
        /// Private constructor to prevent generation of default constructor.
        /// </summary>
        private SpeechSynthesizerHelper()
        {
        }

        static public string GeneratedFilename;

        /// <summary>
        /// Returns temporary file with MP3 for input text, synthetized in call to web s
        /// </summary>
        /// <param name="text">Text to become speech.</param>
        /// <param name="voice">Voice. Default is Joanna in English, US.</param>
        /// <returns></returns>
        static public string GetSpeech(string text, string voice = "")
        {
            // Creates an instance of a speech config with specified subscription key and service region.
            // Replace with your own subscription key and service region (e.g., "westus").
            // var config = SpeechConfig.FromSubscription("YourSubscriptionKey", "westus2");
            var config = SpeechConfig.FromSubscription("c01d27b06a3d4155b335e136717ca4e7", "eastus");

            // Creates a speech synthesizer using the default speaker as audio output.
            using (var synthesizer = new SpeechSynthesizer(config))
            {
                var result = AsyncHelper.RunSync(async () =>
                {
                    return await synthesizer.SpeakTextAsync(text);
                });
                using (result)
                {
                    if (result.Reason == ResultReason.SynthesizingAudioCompleted)
                    {
                        Console.WriteLine($"Speech synthesized to speaker for text [{text}]");

                        GeneratedFilename = Path.Combine(Path.GetTempPath(), DateTime.Now.Ticks.ToString(CultureInfo.InvariantCulture) + "_" + Guid.NewGuid().ToString() + ".m4a");
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

            return GeneratedFilename;
        }
    }
}
