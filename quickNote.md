# key parameters used

`pip3 install python-docx` - imported document package
`pip3 install pyttsx3` - pckage/ library for text to speech

you can choose to include these two packages above with their corresponding version for reusability in a seperate file like *requirement.txt*
python-docx==0.8.11
pyttsx3 == 2.90

When the need to reuse arises simply input the following code to instaall
`pip3 install -r requirement.txt`

To define image sizes use from docx.shared import Inches

To format documents, use aany of the following commands 
document.add_paragraph()
document.add_run()
document.add_picture()
document.add_heading()

