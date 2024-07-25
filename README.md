Step Recorder

Step Recorder is a Python application that captures screenshots on every mouse click and compiles them into a PowerPoint presentation. It features a GUI built with Tkinter and saves screenshots using PyAutoGUI. The generated PowerPoint presentations can be customized with additional text, shapes, and other multimedia.
Features

    Capture screenshots with each mouse click.
    Save screenshots in a designated folder.
    Compile screenshots into a PowerPoint presentation.
    Customize and edit the PowerPoint presentation.
    Export the presentation as a video or webpage.

Installation

    Clone the repository:

    bash

git clone https://github.com/your-username/step-recorder.git
cd step-recorder

Install the required packages:

bash

pip install -r requirements.txt

Run the application:

bash

    python step_recorder.py

Usage

    Create Folder: Click the "Create Folder" button to name the folder where images and the PowerPoint (PPTX file) will be saved. This folder will default to the desktop.

    Start Recording: Click "Start Recording" to begin recording your steps. An image will be captured with each mouse click.

    Stop Recording: Click "Stop Recording" to stop recording your steps. You will be prompted to name the PowerPoint presentation, which will be saved in the created folder.

    View and Edit: Open the folder with your recordings. You will have individual images from each mouse click and a PowerPoint presentation that you can edit and customize.

    Export: The PowerPoint presentation can be exported as a video or webpage using PowerPoint's built-in export features.

Directions

    Click the "Create Folder" button to name the folder where images and the PowerPoint (PPTX file) will be exported to once the recording has finished. This folder will default to the desktop.

    Click "Start Recording" to start recording your steps. An image will be created with each mouse click.

    Click "Stop Recording" to stop the recording of your steps.

    Open the folder with your recording. You will have images from each mouse click and a PowerPoint presentation that you can edit. Name the steps, delete steps if needed, and use any transitions and effects within PowerPoint. Please save once you are finished. You can also export the PPTX as a video file.

    A folder will be created with a PowerPoint of each image on a slide, where you can edit the PowerPoint, adding captions, shapes, text, voice recordings, links to specific sites, etc. You will also have each individual photo image that was created with each mouse click.

    Advanced - Exporting: Export PowerPoint (PPTX) as a webpage (HTML) using the PPTX2HTML tool.

License

This project is licensed under the MIT License - see the LICENSE file for details.
Contributing

Contributions are welcome! Please feel free to submit a Pull Request or open an Issue to improve the project.
Acknowledgements

    PyAutoGUI
    Pillow
    Pynput
    python-pptx

