# PptPolly - Add Amazon Polly voice to a PowerPoint presentation

Yet another spare time piece of code to automate things I do...

## Instructions

- Get code
- Download and install Visual Studio Community Edition.
- Follow instructions to [Configure a computer to develop Office solutions](https://docs.microsoft.com/en-us/visualstudio/vsto/configuring-a-computer-to-develop-office-solutions?view=vs-2019).
- You may need to install the Nugets for AWSSDK.Core and AWSSDK.Polly (should be downloaded automatically, or see below article "Using Amazon Polly from .NET / C#, Get MP3 file")
- After opening the `PptPolly.sln`, open `app.config` and set the values for your AWS `AccessKeyID` and `SecretKey`.
  - Make sure you set the policy AmazonPollyFullAccess for the IAM user.
- Compile and start the project. Go to a PowerPoint slide and enter these lines in the notes
```
This is a test
Matthew
```
- Go to Add-ins tab, and in the PptPolly group, click "Add Voice"
- After a brief webservice call, and audio should be inserted in the slide
- If there is a textbox with title "Caption", the first line appears in the textbox

## TODOs
- Investigate why audio won't playback automatically. But if go to Animation and change to something else and then back to automatic, it will playback.
- Progress during webservice call?
- Button to automate task for all slides at once
- Button to remove all other audio (Slide Show -> Record Slide Show -> Clear Narration on All Slides)
- Migrate to FxCop analyzers: https://aka.ms/fxcopanalyzers

## Troubleshooting

- Running results in exception: `Amazon.Polly.AmazonPollyException: 'Missing Authentication Token'`
  - Confirm that AWS `AccessKeyID` and `SecretKey` in `app.config` are correct.

# References

- Visual Studio Downloads
  - https://visualstudio.microsoft.com/downloads/
- Configure a computer to develop Office solutions
  - https://docs.microsoft.com/en-us/visualstudio/vsto/configuring-a-computer-to-develop-office-solutions?view=vs-2019
- Using Amazon Polly from .NET / C#, Get MP3 file
  - https://chrisbitting.com/2017/04/07/using-amazon-polly-from-net-c-get-mp3-file/
- Using Identity-Based Policies (IAM Policies) for Amazon Polly
  - https://docs.aws.amazon.com/polly/latest/dg/using-identity-based-policies.html
- Understanding and Getting Your Security Credentials
  - https://docs.aws.amazon.com/general/latest/gr/aws-sec-cred-types.html#access-keys-and-secret-access-keys
- Walkthrough: Create your first VSTO Add-in for PowerPoint
  - https://docs.microsoft.com/en-us/visualstudio/vsto/walkthrough-creating-your-first-vsto-add-in-for-powerpoint?view=vs-2017
- Powerpoint add-on to get text in notes in slides and convert it to audio. Doesn't seem to be getting the notes in the slides like it should?
  - https://stackoverflow.com/questions/20975165/powerpoint-add-on-to-get-text-in-notes-in-slides-and-convert-it-to-audio-doesn
