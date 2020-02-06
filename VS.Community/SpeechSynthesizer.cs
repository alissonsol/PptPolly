//----------------------------------------------------------------------
// <copyright file="SpeechSynthesizer.cs" company="Alisson Sol">
//   Code provided "as is", with full rights for any use or change.
// </copyright>
//----------------------------------------------------------------------

using Amazon.Polly;
using Amazon.Polly.Model;
using Amazon.Runtime;
using System;
using System.Globalization;
using System.IO;

namespace PptPolly
{
    internal class SpeechSynthesizer
    {
        // See: https://aws.amazon.com/blogs/security/wheres-my-secret-access-key/
        static string PptPollyAccessKeyID = Properties.Settings.Default.PptPollyAccessKeyID;  // Format example: AKIAIOSFODNN7EXAMPLE
        static string PptPollySecretKey = Properties.Settings.Default.PptPollySecretKey;      // Format example: wJalrXUtnFEMI/K7MDENG/bPxRfiCYEXAMPLEKEY
        static AWSCredentials PptPollyCredentials = new BasicAWSCredentials(PptPollyAccessKeyID, PptPollySecretKey);
        static AmazonPollyClient PollyClient = new AmazonPollyClient(PptPollyCredentials, Amazon.RegionEndpoint.USEast1);

        /// <summary>
        /// Private constructor to prevent generation of default constructor.
        /// </summary>
        private SpeechSynthesizer()
        {
        }

        /// <summary>
        /// Get voice id from string name.
        /// List of voices available: https://docs.aws.amazon.com/polly/latest/dg/voicelist.html
        /// </summary>
        /// <param name="voice">Voice. Default is Joanna in English, US.</param>
        /// <returns></returns>
        static private VoiceId VoiceIdFromString(string voice = "")
        {
            VoiceId voiceId = VoiceId.Joanna; // Default

            switch (voice)
            {
                // English, US
                case "Salli": voiceId = VoiceId.Salli; break;         // Salli, Female
                case "Kimberly": voiceId = VoiceId.Kimberly; break;   // Kimberly, Female
                case "Kendra": voiceId = VoiceId.Kendra; break;       // Kendra, Female
                // case "Joanna": voiceId = VoiceId.Joanna; break;       // Joanna, Female
                case "Ivy": voiceId = VoiceId.Ivy; break;             // Ivy, Female
                case "Matthew": voiceId = VoiceId.Matthew; break;     // Matthew, Male
                case "Justin": voiceId = VoiceId.Justin; break;       // Justin, Male
                case "Joey": voiceId = VoiceId.Joey; break;           // Joey, Male
                // English, British
                case "Emma": voiceId = VoiceId.Emma; break;           // Emma, Female
                case "Amy": voiceId = VoiceId.Amy; break;             // Amy, Female
                case "Brian": voiceId = VoiceId.Brian; break;         // Brian, Female
            }

            return voiceId;
        }

        /// <summary>
        /// Returns temporary file with MP3 for input text, synthetized in call to Amazon Polly.
        /// </summary>
        /// <param name="text">Text to become speech.</param>
        /// <param name="voice">Voice. Default is Joanna in English, US.</param>
        /// <returns></returns>
        static public string GetSpeech(string text, string voice = "")
        {
            SynthesizeSpeechRequest speechRequest = new SynthesizeSpeechRequest();
            speechRequest.Text = text;
            speechRequest.OutputFormat = OutputFormat.Mp3;
            speechRequest.VoiceId = VoiceIdFromString(voice);
            SynthesizeSpeechResponse speech = PollyClient.SynthesizeSpeech(speechRequest);

            string speechFilename = Path.Combine(Path.GetTempPath(), DateTime.Now.Ticks.ToString(CultureInfo.InvariantCulture) + "_" + Guid.NewGuid().ToString() + ".mp3");
            using (var fileStream = File.Create(speechFilename))
            {
                speech.AudioStream.CopyTo(fileStream);
                fileStream.Flush();
            }

            return speechFilename;
        }
    }
}
