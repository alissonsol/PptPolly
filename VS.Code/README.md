# PptPolly - Add Amazon Polly voice to a PowerPoint presentation

Yet another spare time piece of code to automate things I do...

## Instructions

- Git clone the code locally
- Download and install [Visual Studio Code](http://aka.ms/visualstudiocode)
- Configure the computer to develop Office solutions ([link](https://code.visualstudio.com/docs/other/office))
  - Download [Node.js](https://nodejs.org/en/). Make sure to install the LTS (Long Term Support) version.
    - During the "Node.js Setup", select the option to "Automatically install the necessary tools". This will open a command prompt, invoke PowerShell, install Python, Chocolatey, etc. It takes some time, but it will save you time later... but you may be now in Python version path hell.
  - `npm install -g yo bower`
  - `npm install -g generator-office`
  - `yo office`
    - Answers were:
      - Choose a project type: Office Add-in Task Pane project
      - Choose a script type: TypeScript
      - What do you want to name your add-in? PptPolly
      - Which Office client application would you like to support? Powerpoint
  - `npm start`
    - Got error message "[We can't open this add-in from localhost](https://docs.microsoft.com/en-us/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost)" when clicking in "Show Taskpane"?
      - Option 1 in the link above shoudl work: opening administrator command prompt and running the `CheckNetIsolation` command.
  - `code .`
- You may need to install the Nugets for AWSSDK.Core and AWSSDK.Polly (hopefully downloaded automatically, or follow [link](https://chrisbitting.com/2017/04/07/using-amazon-polly-from-net-c-get-mp3-file/)

- After opening the project in Visual Studio Code, open app.config and set the values for your AWS AccessKeyID and SecretKey. Make sure you set the policy AmazonPollyFullAccess for the [IAM user](https://console.aws.amazon.com/iam/home?#)
- Compile and start the project. Go to a PowerPoint slide and enter these lines in the notes
```
This is a test
Matthew
```
- Go to Add-ins tab, and in the PptPolly group, click "Add Voice"
- After a brief webservice call, and audio should be inserted in the slide
- If there is a textbox with title "Caption", the first line appears in the textbox

## TODOs
- Migrate code from CSharp version based on Visual Studio Community to the TypeScript version in VS Code

# References

Visual Studio Downloads
https://visualstudio.microsoft.com/downloads/

Using Amazon Polly from .NET / C#, Get MP3 file
https://chrisbitting.com/2017/04/07/using-amazon-polly-from-net-c-get-mp3-file/

Understanding and Getting Your Security Credentials
https://docs.aws.amazon.com/general/latest/gr/aws-sec-cred-types.html#access-keys-and-secret-access-keys

Walkthrough: Create your first VSTO Add-in for PowerPoint
https://docs.microsoft.com/en-us/visualstudio/vsto/walkthrough-creating-your-first-vsto-add-in-for-powerpoint?view=vs-2017

Powerpoint add-on to get text in notes in slides and convert it to audio. Doesn't seem to be getting the notes in the slides like it should?
https://stackoverflow.com/questions/20975165/powerpoint-add-on-to-get-text-in-notes-in-slides-and-convert-it-to-audio-doesn

Office Add-ins with Visual Studio Code
https://code.visualstudio.com/docs/other/office
